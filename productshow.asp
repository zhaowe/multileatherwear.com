<!-- #include file="conn/conn.asp" -->
<!-- #include file="inc/function.asp" -->
<!--#include file="inc/Char.asp" -->
<%
dim sql 
Set rs= Server.CreateObject ("ADODB.RecordSet")
sql = "select * from Product_Info where Product_ID="&Cint(request.QueryString("Product_ID"))
rs.open sql, conn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>华源皮具</title>
<link href="css/index.css" rel="stylesheet" type="text/css">
<script language=javascript>
<!--

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

</head>
<body  onload="checkwindow()">
<table width="500" height="240" border="0" cellpadding="2" cellspacing="0" bgcolor="#5E0000">
  <tr> 
    <td width="213" rowspan="4"><img src="uploadImages/<%=rs("Product_Big")%>" width="213" height="240" border="0" style="border:1px solid #ffffff"></td>
    <td width="287" height="25" class="text"><%=rs("Product_Name")%></td>
  </tr>
  <tr> 
    <td height="25" class="text"><%=rs("Product_Color")%></td>
  </tr>
  <tr> 
    <td height="25" class="text"><%=rs("Product_Size")%></td>
  </tr>
  <tr> 
    <td valign="top" class="text"><%=rs("Product_Info")%></td>
  </tr>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
%>