var excelSheet;
var folder;

if (WScript.Arguments.Count() > -100) {
    prepareExcel();
    chooseFolder();
    try { 
        main();
    } catch (e) {
        delete excelSheet;
        throw e;
    }
}

function prepareExcel() {
    excel = WScript.CreateObject("Excel.Application");
    excelFileName = chooseFile();
    var book  = excel.Workbooks.Open(excelFileName);
    excelSheet = book.Worksheets.Item(1);
}

function chooseFile() {
    wShell = WScript.CreateObject("WScript.Shell");
    oExec = wShell.Exec("mshta.exe \"about:<input type=file id=FILE><script>FILE.click();new ActiveXObject(\'Scripting.FileSystemObject\').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>\"");
    sFileSelected = oExec.StdOut.ReadLine();
    return sFileSelected;
}

function chooseFolder() {
    objShell  = WScript.CreateObject( "Shell.Application" );
    objFolder = objShell.BrowseForFolder( 0, "Select Folder", 0, "");
    folder = objFolder.Self.Path + "\\";
}


function main() {
    //  оюью, хлемю йнкнмнй лнфмн гюдюрэ гдеяэ
    columnFiles = "L";      //  йнкнмйю ян яяшкйюлх
    columnNumber = "H";     //  йнкнмйю я мнлепнл назъбкемхъ

    for (var row = 2; row <= excelSheet.UsedRange.Rows.Count; row++){
        adNumber = excelSheet.Range(columnNumber + row).Value;
        links = excelSheet.Range(columnFiles + row).Value;
        downloadLine(adNumber, links);
    }

    WScript.Echo("гЮЦПСГЙЮ ЙЮПРХМНЙ ГЮЙНМВХКЮЯЭ!");
}

function downloadLine(adNumber, links) {
    //  оюью, пюгдекхрекэ хлемх тюикю лнфмн гюдюрэ гдеяэ
    fileDelimeter = "-";    //  бнр рср лефдс йюбшвйюлх
    linksArray = links.split(", ");     //  бнр гдеяэ мсфмн гюдюрэ вел пюгдекемш яяшкйх - рнкэйн гюоършлх, хкх гюоършлх я опнаекюлх

    for (var index = 0; index < linksArray.length; index ++) {
        arr = [adNumber, index + 1];
        fileName = arr.join(fileDelimeter);
        //WScript.Echo("\"" + fileName + "\"");
        link = linksArray[index];
        //WScript.Echo("\"" + link + "\"");
        HTTPFileGet(link, folder + fileName + ".jpg");
    }
}

function HTTPFileGet(strFileURL, strFileSave) {
    objXMLHTTP = WScript.CreateObject("MSXML2.XMLHTTP");
    objADOStream = WScript.CreateObject("ADODB.Stream");
    objFSO = WScript.Createobject("Scripting.FileSystemObject");

    objXMLHTTP.Open("GET", strFileURL, false);
    objXMLHTTP.Send();

    if (objXMLHTTP.Status == 200) {
        objADOStream.Open();
        objADOStream.Type = 1;

        objADOStream.Write(objXMLHTTP.ResponseBody);
        objADOStream.Position = 0;

        if (objFSO.FileExists(strFileSave)) {
            objFSO.DeleteFile(strFileSave);
        }

        objADOStream.SaveToFile(strFileSave);
        objADOStream.Close();
    }
}