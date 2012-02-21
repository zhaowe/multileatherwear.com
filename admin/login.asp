<!-- #include file="conn/conn.asp" -->
<!-- #include file="../inc/function.asp" -->
<%

select case request("action")
case "login"  '登录
 call login()
case "logout"  '退出
 call logout()
case else
 call loginPage() '登录
end select 

sub login()  ' ==========================登录================================
msgtitle="用户登录"
Admin_Username=Checkstr(request.form("Admin_Username"))
Admin_Password=Checkstr(Trim(Request.Form("Admin_Password")))
	set rs=server.createobject("adodb.recordset")
	sql="select * from Administrator where Admin_Username='" &Admin_Username& "' and Admin_Password='" &Admin_Password& "'"
	rs.open sql,conn,1,3
	if not rs.eof and not rs.bof then	
			  session("sAdmin_Username")=rs("Admin_Username")
			  session("sAdmin_Password")=rs("Admin_Password")
			   rs("Admin_TimeEnd")=Now()
			   rs("Admin_IP")=Request.ServerVariables("REMOTE_ADDR") 
			   rs.update
			   rs.close
			   set rs=nothing
			  response.Redirect"admin_index.asp"
	else
		response.Redirect"login.asp"		   	   
	end if
	rs.close
	set rs=nothing
end sub

sub logout()  ' ==========================退出================================
			  session("sAdmin_Username")=""
			  session("sAdmin_Password")=""
			  response.Redirect"login.asp"
end sub
sub loginPage()
%>
<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1479" name=GENERATOR>
</HEAD>
<SCRIPT language=javascript>
 function Juge(){
    if (document.myform.Admin_Username.value==""){
       alert ("请输入用户名！");
       document.myform.Admin_Username.focus();
       return(false);
    }

    if (document.myform.Admin_Password.value==""){
       alert ("请输入密码！");
       document.myform.Admin_Password.focus();
       return(false);
    }
}
</SCRIPT>
<BODY bgColor=#799ae1 leftMargin=0 topMargin=150 onload=snow() marginheight="0">
<SCRIPT src="snow.js" type=text/javascript></SCRIPT>
<TABLE width=250 
height=111 border=1 align="center" cellPadding=4 cellSpacing=0 borderColor=#111111 id=AutoNumber1 style="BORDER-COLLAPSE: collapse">
  <TBODY>
  <TR>
    <TD align=middle bgColor=#235acb height=22><FONT 
      color=#ffffff><STRONG>管理员登陆入口</STRONG></FONT> 
  <TR>
    <TD align=middle>
      <FORM action="login.asp" method="post" name="myform" id="myform" onSubmit="return Juge(this)" >
          <strong> 
          <INPUT name=action 
      type=hidden id="action" value="login">
          </strong> 
          <TABLE id=AutoNumber2 style="BORDER-COLLAPSE: collapse" 
      borderColor=#111111 height=100 cellSpacing=0 cellPadding=0 width="100%" 
      border=0>
        <TBODY>
        <TR>
          <TD align=middle height=13>
          <TD height=13>
        <TR>
          <TD align=middle width="32%" height=31><FONT 
            color=#000000>用户名</FONT> 
          <TD width="68%" height=31><INPUT name="Admin_Username" id="Admin_Username" 
            style="BORDER-RIGHT: 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: 1px solid; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-LEFT: 1px solid; PADDING-TOP: 1px; BORDER-BOTTOM: 1px solid" 
            size=15> 
        <TR>
          <TD align=middle width="32%" height=31><FONT color=#000000>密&nbsp; 
            码</FONT> 
          <TD width="68%" height=31><INPUT name="Admin_Password"
            type=password id="Admin_Password" 
            style="BORDER-RIGHT: 1px solid; PADDING-RIGHT: 4px; BORDER-TOP: 1px solid; PADDING-LEFT: 4px; PADDING-BOTTOM: 1px; BORDER-LEFT: 1px solid; PADDING-TOP: 1px; BORDER-BOTTOM: 1px solid" size=15> 
        <TR>
          <TD width="100%" colSpan=2 height=26>
            <P align=center>&nbsp;&nbsp;<INPUT class=buttonface type=submit value=提交 name=intereye> 
            &nbsp;&nbsp;&nbsp; <INPUT class=buttonface type=reset value=重置 name=reset> 
      </P></TR></TBODY></TABLE></FORM></TR></TBODY></TABLE>
	  </BODY>
	  </HTML>
<%
end sub
CloseDatabase
%>