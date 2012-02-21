<!-- #include file="conn/conn.asp" -->
<!--#include file="../inc/function.asp" -->
<!--#include file="../inc/Char.asp" -->
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="style.css" type=text/css rel=stylesheet>
</HEAD>
<SCRIPT language=javascript>
function check()  
{  
	if (document.form.Admin_Username.value=="")
	{   alert("请输入您的性名");
	    document.form.Admin_Username.focus();
		return false;
	}
	if (document.form.Admin_Password.value=="")
	{   alert("请输入您的密码");
	    document.form.Admin_Password.focus();
		return false;
	}
    return true;
}
</SCRIPT>
<link rel="stylesheet" type="text/css" href="style.css">
<body bgcolor="#799AE1" marginheight=0 marginwidth=0 leftmargin=0>
                  <%	
  if request("action")="edit" then
       call SaveEdit()
  else
       call myform()
  end if
  
	
	sub SaveEdit()
      set rs=server.createobject("adodb.recordset")
	  sql="select * from Administrator where Admin_Username='"&Admin_Username&"'"
	  rs.open sql,conn,1,3
	   rs("Admin_Username")=checkstr(request.form("Admin_Username"))
	   rs("Admin_Password")=checkstr(Trim(request.form("Admin_Password")))
	   rs("Admin_TimeStart")=Now()
	   rs.update
	   Admin_Username=""
	   session("sAdmin_Username")=checkstr(request.form("Admin_Username"))
	   Admin_Username= session("sAdmin_Username")
	   call Sysmsg()
	end sub
  
sub myform()
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="24" colspan="3" bgcolor="235ACB">&nbsp;<b><font color=#ffffff>管理员帐号管理</font></b></td>
  </tr>
  <tr bgcolor="#ffffff"> 
    <td width="5%">&nbsp;</td>
    <td width="90%" valign="top" bgcolor="#ffffff">
<table width="100%" border="0" cellspacing="0" cellpadding="3">
        <form name=form onSubmit="return check()" method=post action=Admin_useredit.asp?action=edit>
          <tr> 
            <td width="3%"><img src="images/meeting_memo.gif" width="16" height="16"></td>
            <td width="31%">修改您的名称或密码</td>
            <td width="66%">&nbsp;</td>
          </tr>
          <tr> 
            <td  background="images/linedot.gif" height=1 colspan="3"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>姓名： 
              <input name=Admin_Username class=lineinput id="Admin_Username" value=<%=Admin_Username%> readonly></td>
            <td rowspan="2" valign="bottom"><INPUT  name="submit" type=image src="images/yes.gif" align=absMiddle border=0  title=确定提交></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>密码： 
              <input name=Admin_Password type="password"  class=lineinput id="Admin_Password"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>
<input name=action type=hidden id="action" value="edit"></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td  background="images/linedot.gif" height=1 colspan="3"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </form>
      </table>
    </td>
    <td width="5%" valign="top" bgcolor="#ffffff">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3" height=10 bgcolor="#D6DFF7"></td>
  </tr>
</table>
<% end sub 
		sub Sysmsg()
		%> <br>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TR> 
    <TD bgColor=#235acb height=24>&nbsp;<FONT 
    color=#ffffff><B>修改成功</B></FONT></TD>
  </TR>
</TABLE>
<TABLE cellSpacing=0 cellPadding=5 width="90%" align=center bgColor=#ffffff border=0>
  <TR> 
    <TD align=middle>&nbsp; </TD>
  </TR>
  <TR> 
    <TD>名称或密码修改成功！！！</TD>
  </TR>
  <TR> 
    <TD align=center><a href="javascript:history.go(-1)" >&lt;&lt; 返回上一页</a></TD>
  </TR>
</TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
    <TR> 
      <TD bgColor=#d6dff7 height=10></TD>
    </TR>
  </TBODY>
</TABLE>
<br> 
	  <%end sub %> </BODY></HTML>
<%
CloseDatabase
%>