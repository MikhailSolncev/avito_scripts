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
    //  œ¿ÿ¿, »Ã≈Õ¿  ŒÀŒÕŒ  ÃŒ∆ÕŒ «¿ƒ¿“‹ «ƒ≈—‹
    columnFiles = "L";      //   ŒÀŒÕ ¿ —Œ ——€À ¿Ã»
    columnNumber = "H";     //   ŒÀŒÕ ¿ — ÕŒÃ≈–ŒÃ Œ¡⁄ﬂ¬À≈Õ»ﬂ

    for (var row = 2; row <= excelSheet.UsedRange.Rows.Count; row++){
        linksString = excelSheet.Range(columnFiles + row).Value;
        links = getLinksFromLine(linksString)

        adNumber = excelSheet.Range(columnNumber + row).Value;
        downloadLinks(adNumber, links);
    }

    WScript.Echo("«‡„ÛÁÍ‡ Í‡ÚËÌÓÍ Á‡ÍÓÌ˜ËÎ‡Ò¸!");
}

function getLinksFromLine(linksString) {
    links = [];
    if (linksString == undefined)
        return links;
        
    linksArray = linksString.split(", ");     //  ¬Œ“ «ƒ≈—‹ Õ”∆ÕŒ «¿ƒ¿“‹ ◊≈Ã –¿«ƒ≈À≈Õ€ ——€À » - «¿œﬂ“€Ã», “Œ◊ ¿Ã» »À» “Œ◊ ¿Ã» — «¿œﬂ“Œ…
    for (var index = 0; index < linksArray.length; index ++) {
        link = trim(linksArray[index]);
        if (link.length > 0) links.push(link);
    }
    //WScript.Echo(" ÓÎË˜ÂÒÚ‚Ó ÒÒ˚ÎÓÍ: " + links.length)
    return links;
}

function trim(input) {
    result = input;

    index = result.indexOf(" ", 0);
    while (index > -1) {
        if (index == 0) {
            result = result.substr(1);
        } else if (index < result.length - 1) {
            result = result.substr(0, index).concat(result.substr(index + 1));
        } else if (index == result.length - 1) {
            result = result.substr(0, result.length - 1);
        }
        index = result.indexOf(" ", 0);
    }
    return result;
}

function downloadLinks(adNumber, links) {
    //  œ¿ÿ¿, –¿«ƒ≈À»“≈À‹ »Ã≈Õ» ‘¿…À¿ ÃŒ∆ÕŒ «¿ƒ¿“‹ «ƒ≈—‹
    fileDelimeter = "-";    //  ¬Œ“ “”“ Ã≈∆ƒ”  ¿¬€◊ ¿Ã»
    
    for (var index = 0; index < links.length; index ++) {
        arr = [adNumber, index + 1];
        fileName = arr.join(fileDelimeter);
        link = links[index];
        WScript.Echo("Link: \"" + link + "\"");
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