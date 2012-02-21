<%
Admin_Username= session("sAdmin_Username")
if not isnull(Admin_Username) and Admin_Username<>"" then
	call infomain()
else
	response.Redirect"login.asp"
end if
sub infomain()
%>
<HTML><HEAD><TITLE><%=Admin_Username%>,您好！欢迎进入后台管理系统</TITLE>
<META content="MSHTML 6.00.2800.1479" name=GENERATOR>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="style.css" type=text/css rel=STYLESHEET>
<STYLE>.navPoint {
	FONT-SIZE: 9pt; CURSOR: hand; COLOR: white; FONT-FAMILY: Webdings
}
P {
	FONT-SIZE: 9pt
}
</STYLE>

<SCRIPT>
function switchSysBar(){
	if (switchPoint.innerText==3){
		switchPoint.innerText=4
		document.all("frmTitle").style.display="none"
	}
	else{
		switchPoint.innerText=3
		document.all("frmTitle").style.display=""
	}
}
</SCRIPT>
</HEAD>
<BODY style="MARGIN: 0px" scroll=no>
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD id=frmTitle vAlign=center noWrap align=middle name="frmTitle"><IFRAME 
      id=BoardTitle 
      style="Z-INDEX: 2; VISIBILITY: inherit; WIDTH: 180px; HEIGHT: 100%" 
      name=main src="left.asp" 
frameBorder=0></IFRAME>
      <TD style="WIDTH: 10pt" background=images/t2.gif> 
        <TABLE height="100%" cellSpacing=0 cellPadding=0 border=0>
        <TBODY>
        <TR>
          <TD style="HEIGHT: 100%" onclick=switchSysBar()><SPAN class=navPoint 
            id=switchPoint title=关闭/打开左栏>3</SPAN> </TR></TBODY></TABLE></TD>
    <TD style="WIDTH: 100%"><IFRAME id=frmright 
      style="Z-INDEX: 1; VISIBILITY: inherit; WIDTH: 100%; HEIGHT: 100%" 
      name=BoardList src="right.asp" 
      frameBorder=0></IFRAME></TD></TR></TBODY></TABLE></BODY></HTML>
<%end sub%>