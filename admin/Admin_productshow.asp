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
<BODY bgColor=#799ae1 leftMargin=0 marginwidth="0" marginheight="0"><BR>
<%
 dim rs
 set rs=server.createobject("adodb.recordset")
 rs.open "select * from Product_Info where Product_ID="&cstr(request("Product_ID")),conn,1,1
 %>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
    <TR> 
      <TD bgColor=#235acb height=24>&nbsp;<FONT color=#ffffff><B>查看产品</B></FONT></TD>
    </TR>
  </TBODY>
</TABLE><TABLE cellSpacing=0 cellPadding=0 width="90%" align=center bgColor=#ffffff 
border=0>
<TBODY>
  <TR> 
    <TD align=middle> <TABLE cellSpacing=0 cellPadding=2 width="100%" border=0>
          <TBODY>
            <TR> 
              <TD align=right width="14%">&nbsp;</TD>
              <TD width="33%">&nbsp;</TD>
              <TD width="13%">&nbsp;</TD>
              <TD width="40%">&nbsp;</TD>
            </TR>
            <TR> 
              <TD align=right>产品类别：</TD>
              <TD colspan="3"> 
                <%
			    set rsType=server.createobject("adodb.recordset")
				sql="select * from Product_type where ProductType_ID="&rs("ProductType_ID")
				rsType.open sql,conn,1,1
				if rsType.eof and rsType.bof then
					response.write "没有添加类别"
				else
				response.write""&trim(rsType("Product_Type"))&""
				rsType.close
				end if
				%>
              </TD>
            </TR>
            <TR> 
              <TD align=right>产品名称：</TD>
              <TD colspan="3"><%=trim(rs("Product_Name"))%></TD>
            </TR>
            <TR> 
              <TD align=right>颜 色：</TD>
              <TD colspan="3"><%=trim(rs("Product_Color"))%></TD>
            </TR>
            <TR> 
              <TD align=right>尺 码：</TD>
              <TD colspan="3"><%=trim(rs("Product_Size"))%></TD>
            </TR>
            <TR> 
              <TD align=right>产品小图：</TD>
              <TD colspan="3"> <img src="../uploadImages/<%=trim(rs("Product_Smaill"))%>"></TD>
            </TR>
            <TR> 
              <TD align=right>产品大图：</TD>
              <TD colspan="3"> <img src="../uploadImages/<%=trim(rs("Product_Big"))%>"></TD>
            </TR>
            <TR> 
              <TD></TD>
              <TD colSpan=3></TD>
            </TR>
            <TR valign="top"> 
              <TD align=right>产品简介：</TD>
              <TD colSpan=3><%=rs("Product_Info")%></TD>
            </TR>
            <TR> 
              <TD align=right>&nbsp;</TD>
              <TD align=right>&nbsp;</TD>
              <TD>&nbsp;</TD>
              <TD>&nbsp;</TD>
            </TR>
            <TR> 
              <TD align=right>&nbsp;</TD>
              <TD align=right> [<a href="Admin_addproduct.asp?action=edit&Product_ID=<%=rs("Product_ID")%>">修改产品</a>]&nbsp;&nbsp; 
              </TD>
              <TD>&nbsp; [<a href="Admin_product.asp">返回列表</a>]</TD>
              <TD>&nbsp;</TD>
            </TR>
            <TR> 
              <TD colSpan=4>&nbsp; </TD>
            </TR>
        </TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
  <TR>
    <TD bgColor=#d6dff7 height=10></TD></TR></TBODY></TABLE> </FORM>
        <% 
			rs.Close
			set rs=nothing
		%>
</BODY></HTML>
<%
CloseDatabase
%>