function y2k(number) { return (number < 1000) ? number + 1900 : number; }

     var now = new Date();
     var dd = now.getDate() , mt = now.getMonth() + 1 , yy = y2k(now.getYear()) , weekVal = now.getDay();

        if (weekVal==0)
       msg1="星期日";
    else if (weekVal==1)
        msg1="星期一";
    else if (weekVal==2)
        msg1="星期二";
    else if (weekVal==3)
            msg1="星期三";
    else if (weekVal==4)
            msg1="星期四";
    else if (weekVal==5)
            msg1="星期五";
    else if (weekVal==6)
            msg1="星期六";

   document.write("<table width=70 border=0 cellspacing=1 cellpadding=1><tr bgcolor=#000000><td height=60>");
   document.write("<table width=100% height=58 border=0 bgcolor=#FFFFFF align=center><tr align=center><td>\n");
   document.write('<span STYLE="FONT-FAMILY: 宋体; FONT-SIZE: 9pt;text-decoration: none; color: black">\n');
   document.write(yy+"年"+mt+"月<br>");
   document.write('<font color=#FF0000><b><font face="Arial, Helvetica, sans-serif" size="+1">');
   document.write(dd+"</font></b></font><br>");
   document.write(msg1+"</span></td></tr></table></td></tr></table>");
