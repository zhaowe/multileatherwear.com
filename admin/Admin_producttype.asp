<!-- #include file="conn/conn.asp" -->
<!--#include file="../inc/function.asp" -->
<!--#include file="../inc/Char.asp" -->
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="style.css" type=text/css rel=stylesheet>

<STYLE type=text/css>.style1 {
	COLOR: #ff0000
}
</STYLE>

<META content="MSHTML 6.00.2900.2523" name=GENERATOR></HEAD>
<BODY bgColor=#799ae1 leftMargin=0 marginwidth="0" marginheight="0">
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#235acb">
  <tr> 
    <td height="25" bgcolor="#235acb" align="center"><strong><font color="#FFFFFF">�������</font></strong></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#F3F3F3">
        <tr align="center" bgcolor="#FFFFFF" height="20"> 
          <td width="30%"><strong>��������</strong></td>
          <td width="30%"><strong>��������</strong></td>
          <td width="30%" bgcolor="#FFFFFF"><strong>ȷ������</strong></td>
        </tr>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from Product_type",conn,1,1
		  dim follows
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>��û�з���</font></center>"
		  follows=0
		  else
		  do while not rs.EOF
		  %>
        <form name="form1" method="post" action="Admin_Saveproducttype.asp?action=edit&id=<%=int(rs("ProductType_ID"))%>">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td><input name="Product_type" type="text" id="Product_type" size="20" value="<%=trim(rs("Product_type"))%>"></td>
            <td>
			<input name="Orders" type="text" id="Orders" size="4" value="<%=trim(rs("Orders"))%>">
            </td>
			<td><input type="submit" name="Submit" value="�� ��"> &nbsp; <a href="Admin_Saveproducttype.asp?id=<%=int(rs("ProductType_ID"))%>&action=del" onClick="return confirm('��ȷ������ɾ��������')"><font color=red>ɾ��</font></a> 
            </td>
          </tr>
        </form>
        <%rs.MoveNext
          loop
          follows=rs.RecordCount
          end if%>
      </table></td>
  </tr>
</table>
<br>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#235acb">
  <tr> 
    <td height="25" bgcolor="#235acb" align="center"><font color="#FFFFFF"><strong>��ӷ���</strong></font></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">
	<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#EAEAEA">
        <tr align="center" bgcolor="#FFFFFF" height="20"> 
          <td width="30%"><strong>��������</strong></td>
          <td width="30%"><strong>��������</strong></td>
          <td width="30%"><strong>ȷ������</strong></td>
        </tr>
        <form name="form1" method="post" action="Admin_Saveproducttype.asp?action=add">
          <tr align="center" bgcolor="#FFFFFF"> 
            <td><input name="Product_type" type="text" id="Product_type" size="20"></td>
            <td>
			<input name="Orders" type="text" id="Orders" size="4" value="<%=follows+1%>">
            </td>
            <td><input type="submit" name="Submit" value="�� ��"></td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<br>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#799AE1">
  <tr> 
    <td bgcolor="#235acb" height="25" align="center"><font color="#FFFFFF"><strong>����ע�����˵��</strong></font></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> 
	<table width="100%" border="0" align="center" cellpadding="4" cellspacing="0">
        <tr> 
          <td height="20"><font color="#FF6600">����ע��������Ʋ�Ҫ���зǷ��ַ���<br>
            ������ɾ��������ͬʱ����ɾ�������µ����¡�</font></td>
        </tr>
      </table></td>
  </tr>
</table>
</BODY></HTML>
<%
CloseDatabase
%>