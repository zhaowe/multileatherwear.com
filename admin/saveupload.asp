<!--#include file="upload.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>上传图片</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#799ae1" leftmargin="0" topmargin="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="tableBorder">
  <tr>
    <td align="center" class="forumRow"> 
      <script language="Javascript">
function minipic(smileface)
{
	window.opener.document.myform.Product_Smaill.value=smileface;
}
function pic(smileface)
{
	window.opener.document.myform.Product_Big.value=smileface;
}
</script>
      <%
set upload=new upload_5xSoft
set file=upload.file("file1")
formPath="../uploadImages/"
if file.filesize>100 then
fileExt=lcase(right(file.filename,3))
if fileExt="asp" then
Response.Write"文件类型非法"
end if
end if
randomize
ranNum=int(90000*rnd)+10000
filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
filename_noPath=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
if file.FileSize>0 then 
file.SaveAs Server.mappath(FileName)
end if
%>

图片上传成功,[<a href='uploadCP.asp'>返回继续上传</a>]<br>图片是：<a href=Javascript:minipic("<% =filename_noPath%>");>缩略图</a> &nbsp;&nbsp;&nbsp;&nbsp; <a href=Javascript:pic("<% =filename_noPath%>");>放大图</a>
    </td>
  </tr>
</table>
</body>
</html>