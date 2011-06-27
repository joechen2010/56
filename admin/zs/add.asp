<!--#include file="../inc/syscode.asp" -->

<head>
<title>添加产品信息</title>
<LINK href="images/Admin.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" topmargin="10">
<table width="90%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#6699ff">
  <form method="POST" action="addproducts.asp">
  <tr align="center"> 
    <td height="27" colspan="2" background="images/Left01.gif"><b>添 加 企 业 信 息</b></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="30" align="right">专业选择：</td>
    <td height="30">
<select  
      class=inputText name=type size=1>
                  <option value="广告">广告</option>
                  <option value="法律">法律</option>
                  <option value="经济管理">经济管理</option>
                  <option value="工商管理">工商管理</option>
                  <option value="企业管理">企业管理</option>
                  <option value="市场营销">市场营销</option>
                  <option value="金融">金融</option>
                  <option value="中文">中文</option>
                  <option value="教育">教育</option>
                  <option value="财会">财会</option>
                  <option value="新闻">新闻</option>
                  <option value="文秘">文秘</option>
                  <option value="英语">英语</option>
                  <option value="计算机">计算机</option>
                  <option value="电子商务">电子商务</option>
                  <option value="计算机应用">计算机应用</option>
                  <option value="金融与保险">金融与保险</option>
                 
                  <option value="工业与民用建筑">工业与民用建筑</option>
                  <option value="计算机信息管理">计算机信息管理</option>
                  <option value="物业管理">物业管理</option>
                  <option value="旅游管理">旅游管理</option>
                  <option value="酒店管理">酒店管理</option>
                  <option value="艺术设计">艺术设计</option>
                  <option value="环境艺术设计">环境艺术设计</option>
                  <option value="人力资源管理">人力资源管理</option>
                  <option value="房地产经营管理">房地产经营管理</option>
                  <option value="社区管理">社区管理</option>
				  <option value="机电一体化">机电一体化</option>
				  <option value="酒店管理">酒店管理</option>
				  <option value="工商行政管理">工商行政管理</option>
				  <option value="旅游与酒店管理">旅游与酒店管理</option>
				  <option value="土木工程">土木工程</option>
				  <option value="食品工程">食品工程</option>
				  <option value="电力自动化">电力自动化</option>
				  <option value="幼教">幼教</option>
				  <option value="畜牧兽医">畜牧兽医</option>
				  <option value="行政管理">行政管理</option>
				  <option value="物流管理">物流管理</option>
				  <option value="畜牧">畜牧</option>
				  <option value="英语教育">英语教育</option>
				  <option value="广告设计">广告设计</option>
				  <option value="计算机科学">计算机科学</option>
				  <option value="工程管理">工程管理</option>
				  <option value="电力工程">电力工程</option>
				  <option value="通信工程">通信工程</option>
				  <option value="机械工程">机械工程</option>
				  <option value="材料工程">材料工程</option>
				  <option value="国际金融">国际金融</option>
				  <option value="计算机技术与应用">计算机技术与应用</option>
				  <option value="国际贸易">国际贸易</option>
				  <option value="音乐教育">音乐教育</option>
				  <option value="工艺美术">工艺美术</option>
				  <option value="珠宝玉器营销">珠宝玉器营销</option>
				  <option value="艺术舞蹈">艺术舞蹈</option>
				 
               </select>	
	
	</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="30" width="19%" align="right">姓名：</td>
      <td height="30" width="81%">
      <input name="name" type="text" id="webname" style="border: 1px outset #C0C0C0" size="20">    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
      <td height="28" width="19%" align="right">证书号：</td>
      <td width="81%" height="28"><input name="number" type="text" id="weburl" style="border: 1px outset #C0C0C0" size="20"></td>
  </tr>
    <tr bgcolor="#FFFFFF" align="center"> 
      <td height="30" colspan="2"> 
        <input type="submit" value=" 添 加 "  style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="reset" value=" 重 置 " name="B2" style="color: #FFFFFF; border: 1px outset #808080; background-color: #808080">
    <input type="hidden" name="shibie" value="add">      </td>
  </tr>
  </form>
</table>
</body>
</html>
