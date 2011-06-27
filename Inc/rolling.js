<!--
var version = "other"
browserName = navigator.appName;   
browserVer = parseInt(navigator.appVersion);

if (browserName == "Netscape" && browserVer >= 3) version = "n3";
else if (browserName == "Netscape" && browserVer < 3) version = "n2";
else if (browserName == "Microsoft Internet Explorer" && browserVer >= 4) version = "e4";
else if (browserName == "Microsoft Internet Explorer" && browserVer < 4) version = "e3";

function stopError()
{
	return true;
}
window.onerror = stopError;

function wireOpen()
{
	if (version == "e4")
	{
		document.write("<marquee behavior=scroll direction=up width=100% height=150 scrollamount=1 scrolldelay=70 onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function wireOpen2()
{
	if (version == "e4")
	{
		document.write("<marquee behavior=scroll direction=up width=100% height=560 scrollamount=1 scrolldelay=70 onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function wireClose()
{
	if (version == "e4")
	{
		document.write("</marquee>")
	}
}
//-->

