var mode = 6;

if (!String.prototype.trim) {
  (function() {
    // Вырезаем BOM и неразрывный пробел
    String.prototype.trim = function() {
      return this.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
    };
  })();
}

if (mode == 1) {
	var objFS = WScript.CreateObject("Scripting.FileSystemObject");
	var file = objFS.OpenTextFile("filename.txt", 2, true);
}

if (mode == 2) {
	
	var str = "1486194387_1.jpg, 1486194387_2.jpg, 1486194387_3.jpg, 1486194387_4.jpg, 1486194387_5.jpg, 1486194387_6.jpg, 1486194387_7.jpg";
	str = "1486194387_1.jpg";
	//regexp = /([0-9])*_[0-9].jpg/;
	//arr = str.match("([0-9])*_[0-9].jpg");
	//arr = regexp.exec(str);
	//while (arr != null)
	//	WScript.Echo(arr[0]);
	//
	//arr = str.split("([0-9])*_[0-9].jpg");
	//arr = regexp.split(str);
	arr = str.split(",");
	for (var i = 0; i < arr.length; i++)
		WScript.Echo(arr[i]);
}

if (mode == 3) {
	
	picsFile = "https://i.ibb.co/Pc0Szy8/936605173-2.jpghttps://i.ibb.co/51Kmyd4/1486194387-1.jpghttps://i.ibb.co/d6JpVdH/889588758-1.jpghttps://i.ibb.co/mzxg6sc/889588758-2.jpghttps://i.ibb.co/4JT1C9F/936605173-1.jpghttps://i.ibb.co/nn37jh0/936605173-3.jpg";
	picsString = "1486194387_1.jpg, 1486194387_2.jpg, 1486194387_3.jpg, 1486194387_4.jpg, 1486194387_5.jpg, 1486194387_6.jpg, 1486194387_7.jpg";
	namesArr = picsString.split(",");
		
	for (var num = 0; num < namesArr.length; num++){
		shortName = namesArr[num];
		shortName.trim();
		shortName = shortName.replace("_", "-");
		regstr = "https://i.ibb.co/[\\w]{7}/" + shortName;
		WScript.Echo("\"" + regstr + "\"");
		longNameArr = picsFile.match(regstr);
		if (longNameArr != null)
			WScript.Echo(longNameArr[0]);
	}
	
	
}

if (mode == 4) {
	wShell = WScript.CreateObject("WScript.Shell");
	oExec = wShell.Exec("mshta.exe \"about:<input type=file id=FILE><script>FILE.click();new ActiveXObject(\'Scripting.FileSystemObject\').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>\"");
	sFileSelected = oExec.StdOut.ReadLine();
	WScript.Echo(sFileSelected);
}

if (mode == 5) {
	var dickPic = "1486194387_1.jpg, 1486194387_2.jpg, 1486194387_3.jpg, 1486194387_4.jpg, 1486194387_5.jpg, 1486194387_6.jpg, 1486194387_7.jpg";
	str = "1486194387";
	regexp = ""+ str + "_[0-9].jpg";
	regexp = new RegExp(regexp);
	
	//arr = dickPic.match(regexp);
	arr = regexp.exec(dickPic);
	if (arr == null) {
		WScript.Echo("Arr is null");
	} else {
		for (var i = 0; i < arr.length; i++)
			WScript.Echo(arr[i]);
	}
}

if (mode == 6) {
	var dickPic = "1486194387_1.jpg, 1486194387_2.jpg, 1486194387_3.jpg, 1486194387_4.jpg, 1486194387_5.jpg, 1486194387_6.jpg, 1486194387_7.jpg";
	picArr = dickPic.split(",");
	str = "1486194387";
	regexp = ""+ str + "_[0-9].jpg";
	regexp = new RegExp(regexp, 'y');
	
	//arr = dickPic.match(regexp);
	arr = regexp.exec(dickPic);
	if (arr == null) {
		WScript.Echo("Arr is null");
	} else {
		for (var i = 0; i < arr.length; i++)
			WScript.Echo(arr[i]);
	}
}
