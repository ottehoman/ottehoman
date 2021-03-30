/* SQC Fast Processing Lab specific functions */

/* --- CollectData() --- 
   collects various EBL system data from the Nanosuite app
   and puts them into a text file to be mailed as an attachment.
*/
function CollectData() {
/*	we need a text file. If it does exist, delete it first, then re-create it,
    otherwise, just create it.
*/
    var fso = new ActiveXObject("Scripting.FileSystemObject");
	var tempName = fso.BuildPath(fso.GetSpecialFolder(2), "FPL-eLine-Plus.log");
	if ( fso.FileExists(tempName) ) fso.DeleteFile(tempName);
	var textfile = fso.CreateTextFile(tempName, true);

	// App.WaitMsg("CollectData()", tempName, 5, 1);   // debug only
	
	PL = App.ActivePositionList;
	if (PL) 
	{ 
		// Position list related data
		textfile.WriteLine("PositionList Filename    = " + PL.filename);
		textfile.WriteLine("PositionList ItemIndex   = " + PL.ItemIndex);
		textfile.WriteLine("PositionList Count       = " + PL.Count);
		textfile.WriteLine("PositionList Percent     = " + PL.ItemIndex/PL.Count);
		textfile.WriteLine("PositionList Line No.    = " + PL.GetColData("NO"));
		textfile.WriteLine("PositionList Line ID.    = " + PL.GetColData("ID"));
		textfile.WriteLine("PositionList File        = " + PL.GetColData("File"));
		textfile.WriteLine("PositionList Comment     = " + PL.GetColData("Comment"));
		textfile.WriteLine() ;
		// Column related data
		textfile.WriteLine("Column Aperture size     = " + Column.ApertureSize);
		textfile.WriteLine("Column Aperture X        = " + Column.ApertureX);
		textfile.WriteLine("Column aAperture Y       = " + Column.ApertureY);
		textfile.WriteLine("Column EHT               = " + Column.HighTension);
		textfile.WriteLine("Column Detector          = " + Column.Detector);
		textfile.WriteLine("Column Last Error        = " + Column.LastError);
		textfile.WriteLine("Column Magnification     = " + Column.Magnification);
		textfile.WriteLine("Column Writefield        = " + Column.GetWriteField());
		textfile.WriteLine("Column Stigmator X       = " + Column.StigmatorX);
		textfile.WriteLine("Column Stigmator Y       = " + Column.StigmatorY);
		textfile.WriteLine("Column Working Distance  = " + Column.WorkingDistance);
		textfile.WriteLine();
		// Data file related data
		textfile.WriteLine("GDSII Data File          = " + App.GetVariable("GDSII.DataFile"));
		textfile.WriteLine("GDSII Structure Name     = " + App.GetVariable("Variables.StructureName"));
		textfile.WriteLine();
		// Exposure related data
		textfile.WriteLine("Area Dweltime            = " + App.GetVariable("Variables.Dwelltime"));
		textfile.WriteLine("Exposure Dwell Time      = " + App.GetVariable("Variables.Dwelltime"));
		du = Math.round(1E4*App.GetVariable("Variables.MetricStepX")*App.GetFloatVariable("Exposure.StepSizeU"))/1E4; 
		textfile.WriteLine("Exposure Step Size U     = " + du);
		dv = Math.round(1E4*App.GetVariable("Variables.MetricStepY")*App.GetFloatVariable("Exposure.StepSizeV"))/1E4;
		textfile.WriteLine("Exposure Step Size V     = " + dv);
		textfile.Close();
		}
	// CollectData() returns the name of the file the data was written into.
	return tempName;
}


/* --- sendGmail() --- 
   Collects "To: Subject: Body: Attachment: " from the active positionlist and composes a mail message from that.
   
   Then uses sendGmail.ps1 powershell script in %USERROOT%\Script
   
   sendGMail.ps1 relies on Send-MailMessage cmdlet and requires username/password to be stored as generic credential in Windows, linked to smtp account (currently using smtp.gmail.com)
*/
function sendGmail() {
	/* sends an email from a position list
	  
	   usage: include '-To some.person@some.address.net \
	                   -Subject "A subject" 
	                   -Body "Some text for the body"
	                   -Attach "path_to_a_file.txt"'
	   missing -To will send to no-reply@unsw.edu.au
	   from will be no-reply@unsw.edu.au
	  
	   (c) ojh 2021-03-29 with (help from the internet)
	*/

	//   needs fso file system object to find path to "powershell.exe"
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var wshShell = new ActiveXObject("WScript.Shell");

	var PList = App.ActivePositionList;
	var comment = PList.GetColData("Comment");
	// App.ErrMsg(0,0,comment);							// debug
	// App.WaitMsg("sendGmail()", comment, 5, 1);		// debug

    // CollectData() returns a filename to a text file with collected data
	var attachment = CollectData();
	var options = comment + ' -Attach '+ attachment;
	
	var scriptdir = App.GetSysVariable("VARIABLES.UserRoot") + "\script";
	var ps_cmd = 'powershell -ExecutionPolicy Unrestricted -file "' + scriptdir + '\sendGmail.ps1" ' + options;
	// App.ErrMsg(0,0,ps_cmd);							// debug

	// we create a temporary textfile to log all the output from the commandline (incl. powershell)
	do {
		var tempName = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName());
	} while ( fso.FileExists(tempName) );

	var cmdLine = fso.BuildPath(fso.GetSpecialFolder(1), "cmd.exe") + ' /C '+ ps_cmd + ' > "' + tempName + '"';
	var result = "";
	wshShell.Run(cmdLine, 0, true);

	try {
		var ts = fso.OpenTextFile(tempName, 1, false);
		result = ts.ReadAll();
		ts.Close();
	}
		catch(err) {
		// ignore any errors ...
	}
	// finally, be a good hacker, clean up the mess, leave no traces.
	if ( fso.FileExists(tempName) ) fso.DeleteFile(tempName);
	if ( fso.FileExists(attachment) ) fso.DeleteFile(attachment);	

	App.WaitMsg("sendGmail()", "Mail sent!", 3, 1);

	return result;
}
