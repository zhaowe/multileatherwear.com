<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0053)http://www.footwear-world.com/administrator/right.asp -->
<HTML><HEAD><TITLE>����ҳ��</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>BODY {
	BACKGROUND: #799ae1; MARGIN: 0px; FONT: 9pt ����
}
TABLE {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
TD {
	FONT: 12px ����
}
IMG {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; VERTICAL-ALIGN: bottom; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
A {
	FONT: 12px ����; COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #428eff; TEXT-DECORATION: none
}
.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}
.menu_title {
	
}
.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 8px; CURSOR: hand; COLOR: #215dc6; POSITION: relative; TOP: 2px
}
.menu_title2 {
	CURSOR: hand
}
.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 8px; CURSOR: hand; COLOR: #428eff; POSITION: relative; TOP: 2px
}
</STYLE>

<META content="MSHTML 6.00.2800.1479" name=GENERATOR></HEAD>
<BODY><BR>
<TABLE cellSpacing=0 cellPadding=5 width=600 align=center bgColor=#cccccc 
border=0>
  <TBODY>
  <TR>
    <TD bgColor=#235acb>
      <DIV align=center><FONT color=#ffffff><STRONG><%=session("sAdmin_Username")%>���ã����ɹ������̨����</STRONG></FONT></DIV></TD></TR>
  <TR>
    <TD bgColor=#ffffff>
      <DIV align=center>������ߵ�ѡ��ѡ����Ĳ�������</DIV></TD></TR>
  <TR>
    <TD></TD></TR></TBODY></TABLE><BR>
	<%
	Dim theInstalledObjects(9)
    theInstalledObjects(0) = "Scripting.FileSystemObject"
    theInstalledObjects(1) = "adodb.connection"
    
    theInstalledObjects(2) = "SoftArtisans.FileUp"
    theInstalledObjects(3) = "SoftArtisans.FileManager"
    theInstalledObjects(4) = "JMail.SMTPMail"
    theInstalledObjects(5) = "CDONTS.NewMail"
    theInstalledObjects(6) = "Persits.MailSender"
    theInstalledObjects(7) = "LyfUpload.UploadFile"
    theInstalledObjects(8) = "Persits.Upload.1"
%>
<TABLE cellSpacing=0 borderColorDark=#ffffff cellPadding=5 width=600 
align=center borderColorLight=#dddddd border=0>
  <TBODY>
  <TR bgColor=#235acb>
    <TD class=red_3 colSpan=2><FONT 
      color=#ffffff><STRONG>���������йز���</STRONG></FONT></TD></TR>
  <TR bgColor=#ffffff>
      <TD width=274>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����������������DNS��������IP��ַ��</TD>
      <TD width=306>&nbsp; <%=request.ServerVariables("SERVER_NAME")%></TD>
    </TR>
  <TR bgColor=#ffffff>
    <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������IP��</TD>
      <TD>&nbsp; <%=Request.ServerVariables("LOCAL_ADDR")%></TD>
    </TR>
  <TR bgColor=#ffffff>
    <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������˿ڣ�</TD>
      <TD>&nbsp; <%=request.ServerVariables("SERVER_PORT")%></TD>
    </TR>
  <TR bgColor=#ffffff>
    <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������ʱ�䣺</TD>
      <TD>&nbsp; <%=Now()%></TD>
    </TR>
  <TR bgColor=#ffffff>
    <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;IIS�汾��</TD>
      <TD>&nbsp; <%=request.ServerVariables("SERVER_SOFTWARE")%></TD>
    </TR>
  <TR bgColor=#ffffff>
    <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����������ϵͳ��</TD>
      <TD>&nbsp; <%=Request.ServerVariables("OS")%></TD>
    </TR>
  <TR bgColor=#ffffff>
    <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ű���ʱʱ�䣺</TD>
      <TD>&nbsp; <%=Server.ScriptTimeout%> ��</TD>
    </TR>
  <TR bgColor=#ffffff>
    <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;վ������·����</TD>
      <TD>&nbsp; <%=request.ServerVariables("APPL_PHYSICAL_PATH")%></TD>
    </TR>
  <TR bgColor=#ffffff>
    <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������CPU������</TD>
      <TD>&nbsp; <%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%>��</TD>
    </TR>
  <TR bgColor=#ffffff>
    <TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������������棺</TD>
      <TD>&nbsp; <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></TD>
    </TR></TBODY></TABLE><BR>
<TABLE cellSpacing=0 cellPadding=5 width=600 align=center>
  <TBODY>
  <TR bgColor=#235acb>
    <TD colSpan=2><FONT color=#ffffff><STRONG>���֧���йز���</STRONG></FONT></TD></TR>
            <%
    dim i
    For i=0 to 8
      Response.Write "<TR borderColorLight=#dddddd bgColor=#ffffff borderColorDark=#ffffff><TD width=271>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & theInstalledObjects(i) & ""
	  Response.Write "</font></td>"
      If Not IsObjInstalled(theInstalledObjects(i)) Then 
        Response.Write "<TD width=307>&nbsp; <b>��</b> ����֧�֣� </td>"
      Else
        Response.Write "<TD width=307>&nbsp; <b>��</b> ��֧�֣� </td>"
      End If
      Response.Write "</TR>" & vbCrLf
    Next
    %>
  <TR borderColorLight=#dddddd bgColor=#d6dff7 borderColorDark=#ffffff>
    <TD colSpan=2 height=10></TD></TR></TBODY></TABLE><BR></BODY></HTML>
    <%
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>