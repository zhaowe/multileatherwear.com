<!--#include file="upload.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ϴ�ͼƬ</title>
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
Response.Write"�ļ����ͷǷ�"
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

ͼƬ�ϴ��ɹ�,[<a href='uploadCP.asp'>���ؼ����ϴ�</a>]<br>ͼƬ�ǣ�<a href=Javascript:minipic("<% =filename_noPath%>");>����ͼ</a> &nbsp;&nbsp;&nbsp;&nbsp; <a href=Javascript:pic("<% =filename_noPath%>");>�Ŵ�ͼ</a>
    </td>
  </tr>
</table>
</body>
</html>