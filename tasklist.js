
  if ( WScript.FullName.toUpperCase().indexOf("CSCRIPT.EXE")==-1 ) {
    WScript.Echo("Please run with CScript.exe explicitly or change the default script host to CScript.exe")
    WScript.Quit();
  }

  // > wmic process where processid=19336 get processid,name
  // > wmic process where "name like 'iexplore%'" get processid,name

  wmiQuery = "SELECT Name, CreationDate, ProcessID, CommandLine "
           + " FROM Win32_Process WHERE Name IS NOT NULL ";

  argv = WScript.Arguments;
  if (argv.length > 0) {
    if ( isNaN(argv(0)) ) {
      // wmiQuery = wmiQuery + " AND Name LIKE '" + argv(0) + "'";
      var pid = getPidByWindowTitle(argv(0));
      if (pid==-1) WScript.Quit();
      wmiQuery = wmiQuery + " AND ProcessID=" + pid;
    } else {
      wmiQuery = wmiQuery + " AND PID=" + argv(0);
    }
  }

  //** The WMI Query Language does not support sorting/order by.    **//
  //** You can use a disconnected recordset to sort. See this link: **//
  //** http://technet.microsoft.com/library/ee176578.aspx           **//

  objWMIService = GetObject("winmgmts:\\\\.\\root\\cimv2");
  colProcessList = objWMIService.ExecQuery(wmiQuery);
  for (var objProcess=new Enumerator(colProcessList); !objProcess.atEnd(); objProcess.moveNext()) {
    WScript.Echo( objProcess.item().Name + " [" + objProcess.item().ProcessID + "]"
      + " [" + WMIDateStringToDate(objProcess.item().CreationDate) + "] "
      + objProcess.item().CommandLine
    );
  }

  WScript.Quit()

  //////////////////////////////////////////////////////////////////////

  function WMIDateStringToDate(wmiDateStr) {
    // http://blogs.technet.com/b/heyscriptingguy/archive/2005/07/20/how-can-i-determine-the-date-and-time-that-a-process-started.aspx
    var dateStr, dateVal;
    try {
      dateStr = wmiDateStr.substr(0, 4) + "-"
              + wmiDateStr.substr(4, 2) + "-"
              + wmiDateStr.substr(6, 2) + " "
              + wmiDateStr.substr(8, 2) + ":"
              + wmiDateStr.substr(10, 2) + ":"
              + wmiDateStr.substr(12, 2);
      return(dateStr);

      dateVal = new Date(
                  wmiDateStr.substr(0, 4),
                  parseInt(wmiDateStr.substr(4, 2)) - 1,
                  wmiDateStr.substr(6, 2),
                  wmiDateStr.substr(8, 2),
                  wmiDateStr.substr(10, 2),
                  wmiDateStr.substr(12, 2)
                );
      return(dateVal.toString());
    } catch(e) {}
    return;
  }

  //////////////////////////////////////////////////////////////////////

  function getPidByWindowTitle(windowTitle) {
    var cmdln = 'tasklist /NH /FO CSV /FI "WINDOWTITLE eq ' + windowTitle + '"';
    var output = getCmdLnStdOut(cmdln);
    var rows = csvToArray(output);
    if ( rows.length > 0) {
      cols = rows[0];
      if ( cols.length > 2 && !isNaN(cols[1]) ) {
        return parseInt(cols[1]);
      }
    }
    return -1;
  }

  //////////////////////////////////////////////////////////////////////

  function csvToArray(text) {
    // [ https://stackoverflow.com/questions/8493195/how-can-i-parse-a-csv-string-with-javascript-which-contains-comma-in-data ]
    var p='', cols=[''], rows=[cols], i=0, r=0, s=true, l;
    for (var c=0; c<text.length; c++) {
      l = text.substr(c,1);
      if ('"' === l) {
        if (s && l===p) cols[i]+=l;
        s = !s;
      } else if (','===l && s) {
        l = cols[++i] = '';
      } else if ('\n' === l && s) {
        if ('\r' === p) cols[i] = cols[i].slice(0, -1);
        col = rows[++r] = [l = '']; i = 0;
      } else cols[i] += l;
      p = l;
    }
    return rows;
  }

  //////////////////////////////////////////////////////////////////////

  function getErrCode(err) {
    var errNumber = err.number;
    // get the winerror-style representation of the hex value
    if ( errNumber < 0) errNumber += 0xFFFFFFFF + 1;
    return "0x" + errNumber.toString(16); 
  }

  function getCmdLnStdOut(cmdln) {
    try {
      var responseText="", objExec=new ActiveXObject("WScript.Shell").Exec(cmdln);
      while (! objExec.StdOut.AtEndOfStream) {
        responseText += objExec.StdOut.ReadLine() + "\n";
      } objExec=null;
      return responseText;
    } catch(err) {
      return "<ERR/>WshShell.Exec# " + cmdln + "\nErr# " + getErrCode(err) + ": " + err.description;
    }
    /*
      var tmpFile = "B:\\cmdln.out.txt";
      var cmdln = 'cmd.exe /c " ' + cmdln + ' >"' + tmpFile + '" 2>&1"';
      WScript.CreateObject("WScript.Shell").Run(cmdln, 1, true);
      return WScript.CreateObject("Scripting.FileSystemObject").OpenTextFile(tmpFile, 1).ReadAll();
    */
  }

