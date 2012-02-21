<!-- #include file="conn/conn.asp" -->
<!--#include file="../inc/function.asp" -->
<!--#include file="../inc/Char.asp" -->
<%
if isnull(Admin_Username) and Admin_Username="" then 
		response.Redirect"login.asp"		   	   
end if
%>
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="style.css" type=text/css rel=stylesheet>

<STYLE type=text/css>.style1 {
	COLOR: #ff0000
}
</STYLE>

<META content="MSHTML 6.00.2900.2523" name=GENERATOR></HEAD>
<SCRIPT language=JavaScript>
function Juge(myform)
{
    if (document.myform.GuestBook_Revert.value==""){
     alert("请输入回复内容！")
     document.myform.GuestBook_Revert.focus();
     return(false);
    }
}
</SCRIPT>

<BODY bgColor=#799ae1 leftMargin=0 marginwidth="0" marginheight="0"><BR>
<%
 dim rs
 set rs=server.createobject("adodb.recordset")
 rs.open "select * from GuestBook where GuestBook_ID="&cstr(request("GuestBook_ID")),conn,1,1
 GuestBook_ID=rs("GuestBook_ID")
 %>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
    <TR> 
      <TD bgColor=#235acb height=24>&nbsp;<FONT color=#ffffff><B>客户留言查看</B></FONT></TD>
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
              <TD align=right>姓　名：</TD>
              <TD colspan="3"><%=trim(rs("Name"))%> </TD>
            </TR>
            <TR> 
              <TD align=right>公　司：</TD>
              <TD colspan="3"><%=trim(rs("Inc"))%></TD>
            </TR>
            <TR> 
              <TD align=right>电　话：</TD>
              <TD colspan="3"><%=trim(rs("Tel"))%></TD>
            </TR>
            <TR> 
              <TD align=right>邮　箱：</TD>
              <TD colspan="3"><%=trim(rs("Email"))%></TD>
            </TR>
            <TR> 
              <TD align=right>地　址：</TD>
              <TD colspan="3"><%=trim(rs("Add"))%></TD>
            </TR>
            <TR> 
              <TD></TD>
              <TD colSpan=3></TD>
            </TR>
            <TR valign="top"> 
              <TD align=right>留言内容：</TD>
              <TD colSpan=3><%=HTMLEncode(rs("Info"))%></TD>
            </TR>
            <!--             <TR valign="top"> 
              <TD align=right>&nbsp;</TD>
              <TD colSpan=3>
                <%
						GuestBook_ID=rs("GuestBook_ID")
						dim rsRe
						dim sqlRe
						Set rsRe= Server.CreateObject ("ADODB.RecordSet")
						sqlRe = "select * from GuestBook_Revert where GuestBook_ID="&GuestBook_ID&" order by Revert_Time Desc"
						rsRe.open sqlRe, conn,1,3
						do while not rsRe.eof
						%>
                <font color="#000000">回复：<%=HTMLEncode(rsRe("GuestBook_Revert"))%></font><br> 
                <%
						rsRe.movenext 
						loop 
						rsRe.close
						set rsRe=nothing
						%>
              </TD>
            </TR> -->
            <TR> 
              <TD align=right>&nbsp;</TD>
              <TD align=right>&nbsp;</TD>
              <TD>&nbsp;</TD>
              <TD>&nbsp;</TD>
            </TR>
            <TR> 
              <TD align=right>&nbsp;</TD>
              <TD align=right> 
                <!-- &nbsp; [<a href="Admin_guestbookshow.asp?action=revert&GuestBook_ID=<%=GuestBook_ID%>">回复留言</a>] -->
              </TD>
              <TD>&nbsp; [<a href="Admin_guestbook.asp">返回列表</a>]</TD>
              <TD>&nbsp;</TD>
            </TR>
            <TR> 
              <TD colSpan=4>&nbsp; </TD>
            </TR>
        </TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
  <TR>
    <TD bgColor=#d6dff7 height=10></TD></TR></TBODY></TABLE>
        <% 
			rs.Close
			set rs=nothing
		%>                  <%	
  select case request("action")
  case "add" 
       call SaveAdd()
  case "revert" 
       call myform()
  end select 
	
	sub SaveAdd()
      set rs=server.createobject("adodb.recordset")
	  sql="select * from GuestBook_Revert where (Revert_ID is null)"
	  rs.open sql,conn,1,3
  	   rs.addnew
	   rs("GuestBook_ID")=trim(request.form("GuestBook_ID"))
	   rs("GuestBook_Revert")=request.form("GuestBook_Revert")
	   rs("Revert_Time")=Now()
	   rs.update
		response.Redirect"Admin_guestbookshow.asp?GuestBook_ID="&request.form("GuestBook_ID")&""
	end sub
  
sub myform()
%>
<br>
<br>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
    <TR> 
      <TD bgColor=#235acb height=24>&nbsp;<FONT color=#ffffff><B>客户留言查看</B></FONT></TD>
    </TR>
  </TBODY>
</TABLE>

  <TABLE cellSpacing=0 cellPadding=0 width="90%" align=center bgColor=#ffffff 
border=0>
  <form name="myform" method="post" action="Admin_guestbookshow.asp" onSubmit="return Juge(this)">
    <TBODY>
      <TR> 
        <TD align=middle> <TABLE cellSpacing=0 cellPadding=2 width="100%" border=0>
            <TBODY>
              <TR> 
                <TD align=right width="14%">&nbsp;</TD>
                <TD width="33%"> <input name="action" type="hidden" id="action" value="add">
                  <input name="GuestBook_ID" type="hidden" id="GuestBook_ID" value="<%=GuestBook_ID%>"></TD>
                <TD width="13%">&nbsp;</TD>
                <TD width="40%">&nbsp;</TD>
              </TR>
              <TR> 
                <TD></TD>
                <TD colSpan=3></TD>
              </TR>
              <TR valign="top"> 
                <TD align=right>回复内容：</TD>
                <TD colSpan=3><TEXTAREA name="GuestBook_Revert" cols=83 rows=10 id="GuestBook_Revert"><% if isedit then response.write rse("Product_Info") end if %></TEXTAREA></TD>
              </TR>
              <TR> 
                <TD align=right>&nbsp;</TD>
                <TD align=right>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
              </TR>
              <TR> 
                <TD align=right>&nbsp;</TD>
                <TD align=right><INPUT class=buttonface type=submit value=确认回复 name=bi> 
                  &nbsp;&nbsp; </TD>
                <TD>&nbsp; <INPUT class=buttonface type=reset value="清除 ^.^" name=Submit2> 
                </TD>
                <TD>&nbsp;</TD>
              </TR>
              <TR> 
                <TD colSpan=4>&nbsp; </TD>
              </TR>
          </TABLE></TD>
      </TR>
    </TBODY></form>
  </TABLE>

<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
    <TR> 
      <TD bgColor=#d6dff7 height=10></TD>
    </TR>
  </TBODY>
</TABLE>
</BODY></HTML>
<%
end sub
CloseDatabase
%>