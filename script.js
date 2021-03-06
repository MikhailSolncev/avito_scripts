var excel;
var picsFile = "";

if (WScript.Arguments.Count() > -100) {
	prepare();
	
	try { 
		main();
	} catch (e) {
		delete excel;
		throw e;
	}
}

function prepare() {
	
	excel = WScript.CreateObject("Excel.Application");
	objFS = WScript.CreateObject("Scripting.FileSystemObject");
	excelFileName = chooseFile();
	
	textFileName = chooseFile();
	fileObject = objFS.OpenTextFile(textFileName);
	while (!fileObject.AtEndOfStream) {
		picsFile += fileObject.ReadLine();
	}
	fileObject.Close();
}

function chooseFile() {
	wShell = WScript.CreateObject("WScript.Shell");
	oExec = wShell.Exec("mshta.exe \"about:<input type=file id=FILE><script>FILE.click();new ActiveXObject(\'Scripting.FileSystemObject\').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>\"");
	sFileSelected = oExec.StdOut.ReadLine();
	return sFileSelected;
}

function main() {
	var book  = excel.Workbooks.Open(excelFileName);
	
	var sheet = book.Worksheets.Item(1);

	for (var row = 2; row <= sheet.UsedRange.Rows.Count; row++){
		adNumber = sheet.Range("H" + row).Value;
		result = "";
		
		for (var i = 1; i < 12; i++) {
			//regstr = "https://i.ibb.co/[\\w]{7}/" + adNumber + "-" + i + ".jpg";
			regstr = "http://ipic.su/img/img4/fs/" + adNumber + "_" + i + ".[0-9]{1,}.jpg";
			longNameArr = picsFile.match(regstr);
			if (longNameArr != null) 
				result = "" + result + "###" + longNameArr[0];
		}
		
		result = result.substr(3);
		for (var i = 0; i<10; i++)
			result = result.replace("###", ", ");
		
		sheet.Range("J" + row).Value = result;
	}
	
	book.Close(true);
	
	WScript.Sleep(2000);
	
	WScript.Echo("Excel file modification complete");
}