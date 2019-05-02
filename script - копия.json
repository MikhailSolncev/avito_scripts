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
	//currentDir = objFS.GetParentFolderName(WScript.ScriptFullName);
	//excelFileName = currentDir + "\\" + WScript.Arguments.Item(0);
	excelFileName = chooseFile();
	
	//textFileName = currentDir + "\\" + WScript.Arguments.Item(1);
	textFileName = chooseFile();
	//textFileName = WScript.Arguments.Item(1);
	//WScript.Echo(textFileName);
	fileObject = objFS.OpenTextFile(textFileName);
	while (!fileObject.AtEndOfStream)
		//picsFile += fileObject.ReadLine() + "\n\r";
		picsFile += fileObject.ReadLine();
	fileObject.Close();
	//WScript.Echo("\"" + picsFile + "\"");
	
	if (!String.prototype.trim) {
	  (function() {
		// Вырезаем BOM и неразрывный пробел
		String.prototype.trim = function() {
		  return this.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
		};
	  })();
	}
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
		picsString = sheet.Range("J" + row).Value;
		namesArr = picsString.split(",");
		
		for (var num = 0; num < namesArr.length; num++){
			oldShortName = namesArr[num];
			oldShortName = oldShortName.replace(" ", "");
			
			shortName = oldShortName.replace("_", "-");
			regstr = "https://i.ibb.co/[\\w]{7}/" + shortName;
			longNameArr = picsFile.match(regstr);
			if (longNameArr != null) {
				picsString = picsString.replace(oldShortName, longNameArr[0]);
			} else {
				WScript.Echo("Cant match \"" + shortName + "\"");
			}
		}
		sheet.Range("J" + row).Value = picsString;
	}
	
	book.Close(true);
	
	WScript.Sleep(2000);
	
	WScript.Echo("Excel file modification complete");
}