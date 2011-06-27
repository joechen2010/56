<!--#include file = "../Inc/Syscode.asp"-->

<html>
<head>
</head>

<body>

  <form action="gong_AddOK.asp" method="post" name="form">
	
  <table width="650" align=center cellspacing=3>
      <tr> 
      <td width="71" align="right">标题：</td>
      <td width="564"> <input name="bt" type="text"></td>
    </tr>
    <tr> 
      <td width="71" align="right">公告内容：</td>
      <td width="564"> <input name="Content" type="hidden">
        <iframe ID="eWebEditor1" src="../Editor/ewebeditor.asp?id=Content&style=" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe>      </td>
    </tr>
  </table>
	<p align=center>
	  <input type="submit" name="Submit" value="提  交">
	  <input type=reset name=btnReset value=" 重 填 "></p>
</form>


</body>
</html>
