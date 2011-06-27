
var siteid="";
siteid="40880";

var soft="";
soft=navigator.appName;
if (soft != "")
	soft=escape(soft);

var color="";
if (soft != "Netscape") 
	color=screen.colorDepth;
else 
	color=screen.pixelDepth;

var resolve="";
resolve=screen.width+"x"+screen.height;
resolve=escape(resolve);

var title="";
title=escape(document.title);

var fromurl="";
try {fromurl=top.document.referrer;} catch(err) {fromurl="";} finally {fromurl=(fromurl=="")?document.referrer:fromurl;}
fromurl=escape(fromurl);

var resource="";
resource=document.URL;
if (resource != "")
	resource=escape(resource);

var clientsys="";
clientsys=navigator.userAgent;
if (clientsys != "")
	clientsys=escape(clientsys);

r="?id="+siteid+"&referer="+fromurl+"&resolve="+resolve+"&navigator="+soft+"&color="+color+"&title="+title+"&resource="+resource+"&clientsys="+clientsys;

function getCookieVal(offset) {
	var endstr=document.cookie.indexOf(";",offset);
	if( endstr==-1 )
		endstr=document.cookie.length;
	return document.cookie.substring(offset,endstr);
}

function GetCookie(name) {
	var arg = name + "=";
	var alen = arg.length;
	var clen = document.cookie.length;
	var i = 0;
	var j;
	while( i<clen ) {
		j = i + alen;
		if( document.cookie.substring(i,j)==arg )
			return getCookieVal(j);
		i = document.cookie.indexOf(" ",i)+1;
		if(i==0)
			break;
	}
	return null;
}

var newuser="";
function flux_stat() 
{
	var expireDate = new Date();
	var expiretime = 93312000;
	var c_value = "0.34072200 11536389261130350996";
	expireDate.setTime (expireDate.getTime() + expiretime);
	document.cookie = "flux_stat_user=" + c_value + "; expires=" + expireDate.toGMTString() +  "; ";
	return c_value;
}
offset = document.cookie.indexOf("flux_stat_user");
if (offset == -1) 
{
	newuser = 10;
	coockie_value = flux_stat();
}
else
{
	newuser = 20;
	coockie_value = GetCookie("flux_stat_user");
}

r = r+"&flux_stat_user="+coockie_value; 
r = r+"&flux_new_user="+newuser;

document.open();

document.write("<script language=\"JavaScript\" type=\"text/javascript\" src=\"http://211.100.21.4/info.cnt"+r+"\"></script>");
document.write( "<a href=\"http://new.50bang.com/report_out.htm?id=40880\" target=\"_blank\"><img src=\"http://ww.50bang.com/i/1.gif\" width=\"20\" height=\"20\" border=\"0\" alt=\"武林榜免费流量统计系统\"></a>" );
document.close();
