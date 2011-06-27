document.open();
document.write('<a href=http://report2.itsun.com/stats/index.jsp?uuid=1634743 target=_blank title=ITSUN-中国最专业的免费流量统计-IP:124.90.83.13 style=font-size:9pt;>[ITSUN统计]</a>');
var cookieString = new String(document.cookie); 
var cookieHeader = "Cookie_itsun_1634743=" ;
var beginPosition = cookieString.indexOf(cookieHeader) ;
if (beginPosition == -1){
var Then = new Date() ;
var hours = Then.getHours();
var minutes = Then.getMinutes();
var seconds = Then.getSeconds();
var lefttime = 1000 * ( 86400 - hours*3600 - minutes*60 - seconds);
Then.setTime(Then.getTime() + lefttime );
document.cookie = "Cookie_itsun_1634743=y;expires="+ Then.toGMTString() +"; path=/";
var fromr=top.document.referrer;
fromr=(fromr=="")?document.referrer:fromr;
var c_page=top.location.href;
c_page = c_page =="" ? location.href : c_page;
var query = '&c_page=' + escape(c_page) + '&refer1=' + escape(fromr) + '&c_screen=' + screen.width+ 'x'+screen.height;
document.write('<IFRAME width="0" height="0" marginwidth="0" marginheight="0" frameborder="0" src="http://222.73.254.62\:8888/count2.php?uuid=1634743&ip1=124.90.83.13'+ query + '"></iframe>');
}
document.close();

