document.write("<link href=http://www.mainone.com/searchbar/css.css rel=stylesheet type=text/css> ")
document.write("<script language=javascript src=http://www.mainone.com/searchbar/mainonefunction.js></script>");
document.write("<table width=160 border=0 cellspacing=0 cellpadding=0>")
document.write("<form name=mainonesearchForm method=post onsubmit=mainoneSearch() target=_blank>")
document.write("  <tr> ")
document.write("    <td class=searchbar width=124>&nbsp;<input type=text maxLength=20 name=keyword class=shortText><input type=hidden name=subflag value=true></td>")
document.write("    <td class=searchbar width=42><input type=image src=http://www.mainone.com/searchbar/images/search.gif width=37 height=18 name=sub ></td>")
document.write("  </tr>")
document.write("  <tr> ")
document.write("    <td class=searchbar><input name=For id=mainonecompany value=company type=radio checked><img src=http://www.mainone.com/searchbar/images/qy.gif><input name=For id=mainoneproduct value=product type=radio ><img src=http://www.mainone.com/searchbar/images/cp.gif></td>")
document.write("    <td class=searchbar><input type=image src=http://www.mainone.com/searchbar/images/through.gif name=though ></td>")
document.write("  </tr>")
document.write("</form>")
document.write("</table>")

