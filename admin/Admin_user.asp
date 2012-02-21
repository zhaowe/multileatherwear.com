<!--#include file="conn/conn.asp" -->
<!--#include file="../inc/function.asp" -->

<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2800.1479" name=GENERATOR></HEAD>
<BODY bgColor=#799ae1 leftMargin=0 marginwidth="0" marginheight="0">
<SCRIPT language=JavaScript>
function cha(){
if(selform.Key.value==""){
alert("请输入关键字！");
selform.Key.focus();
return (false);
}
}
</SCRIPT>
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
<SCRIPT language=javascript>
function show_set0(mylink0)
{
  window.open(mylink0,'new0','top=110,left=280,width=650,height=290,scrollbars=no')
}
</SCRIPT>

<SCRIPT language=javascript>
function show_set(mylink)
{
  window.open(mylink,'new','top=110,left=280,width=240,height=120,scrollbars=no')
}
</SCRIPT>

<SCRIPT LANGUAGE=javascript>
<!--
function SelectAll() {
	for (var i=0;i<document.selform.selCateID.length;i++) {
		var e=document.selform.selCateID[i];
		e.checked=!e.checked;
	}
}
//-->
</script>
<% 
dim Type_ID
dim sql,rs
if not isnull(Admin_Username) and Admin_Username<>"" then 
  select case request("action")
  case "add"
  	   call myform()
       call add()
  case "addUser"
  	   call myform()
       call addUser()
  case "del"
       call delCate()
  case else
       call myform() 
  end select
else
		response.Redirect"login.asp"		   	   
end if
  
  sub delCate()
		msgtitle="删除用户"
        conn.execute("delete from Administrator where Admin_ID in ("&Request.Form("selCateID")&")")
	    msginfo="删除成功！<br><a href=""Admin_User.asp"">返回</a><br><a href=""../"">返回首页</a>"
        call Sysmsg(msgtitle,msginfo) 
  end sub
  
sub addUser()  ' ==========================新用户注册================================
     set rs=server.createobject("adodb.recordset")
	 sql="select * from Administrator where Admin_Username='"&request.form("Admin_Username")&"'"
      rs.open sql,conn,1,3
	  msgtitle="用户注册"
	  if not (rs.eof and rs.bof) then		
		msginfo="<br>对不起，您输入的用户名已经被注册，请重新输入。"
	  else
	  rs.addnew	  
      rs("Admin_Username")=Checkstr(request.form("Admin_Username"))
	  rs("Admin_Password")=Checkstr(Trim(Request.Form("Admin_Password")))
	  rs("Admin_TimeStart")=Now()
	  rs("Admin_TimeEnd")=Now()
	  rs("Admin_IP")=Request.ServerVariables("REMOTE_ADDR")
	  rs.update
	  msginfo="添加成功！<br><a href=""Admin_User.asp"">返回</a><br><a href=""../"">返回首页</a>"
	  end if
	  rs.close
	  set rs=nothing
	  call Sysmsg(msgtitle,msginfo)
	  response.Redirect"Admin_user.asp"
end sub

  %> <% sub myform() %> 
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
  <TR>
      <TD bgColor=#235acb height=24>&nbsp;<FONT 
    color=#ffffff><B>管理员帐号管理　[<a href="Admin_user.asp?action=add">添加</a>]</B></FONT></TD>
    </TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
		<form name="selform" method="post" action="Admin_user.asp" > 
  <TBODY>
  <TR>
    <TD align=middle bgColor=#ffffff>
      <TABLE cellSpacing=0 cellPadding=3 width="98%" border=0>
        <TBODY>
        <TR>
          <TD width="5%"><IMG src="images/search.gif"></TD>
          <TD width="95%">&nbsp;关键字: 
                  <INPUT name=Key class=lineInput id="Key"> 
            &nbsp;&nbsp; 按名称查询&nbsp; <INPUT type=image height=18 width=43 
            src="images/search1.gif" align=absMiddle border=0 name=image onClick="return cha()"> 
          </TD></TR></TBODY></TABLE>
      <CENTER>
      <TABLE cellSpacing=0 cellPadding=4 width="95%" border=0>
            <TBODY>
              <TR> 
                <TD background=images/linedot.gif colSpan=6 height=1></TD>
              </TR>
              <TR height=25> 
                <TD width="90" height="25" align=left><FONT color=#ff6600><B>ID号</B></FONT></TD>
                <TD height="25" align=middle><FONT color=#ff6600><B>用户名</B></FONT></TD>
                <TD align=middle><FONT color=#ff6600><B>注册或修改时间</B></FONT></TD>
                <TD align=middle><FONT color=#ff6600><B>最后登录时间</B></FONT></TD>
                <TD height="25" align=middle><FONT color=#ff6600><B>最后登录IP</B></FONT></TD>
                <TD width=40 height="25" align=middle><input name=action type=hidden id="action" value="del">
                  <INPUT class=buttonface type=submit value=删除></TD>
              </TR>
              <TR> 
                <TD background=images/linedot.gif colSpan=6 height=1></TD>
              </TR>
              <%
				   dim sql 
				   dim rs 
				   dim totalPut
				   dim CurrentPage,MaxPerPage
				   dim TotalPages
				   dim i,j    
				   dim errmsg,howmanyfields
				   MaxPerPage=12
				Set rs= Server.CreateObject ("ADODB.RecordSet")
				if request("Key")="" then
				sql = "select * from Administrator order by Admin_ID Desc"
				else
				sql = "select * from Administrator where Admin_Username like '%"&Trim(Request("Key"))&"%' order by Admin_ID Desc"
				end if 
				rs.open sql, conn,1,3
				'----------------------------------
				if rs.eof then 
				response.write "<tr><td></td><td>还没有记录呢！！！</td><tr>"
				response.end
				end if
				rs.MoveFirst  '到第一条记录
				rs.pagesize=MaxPerPage  '定义每页记录数
				howmanyfields=rs.Fields.Count-1  
				
				If trim(Request("Page"))<>"" then
					CurrentPage= CLng(request("Page")) 
					If CurrentPage> rs.PageCount then 
						CurrentPage = rs.PageCount 
					End If 
				Else 
					CurrentPage= 1 
				End If 
				
				totalPut=rs.recordcount 'totalput=总记录数
				if CurrentPage<>1 then 
					if (currentPage-1)*MaxPerPage < totalPut then 
						rs.move(currentPage-1)*MaxPerPage			
					end if 
				end if
				
				'计算总页数
				totalpages=rs.pagecount
				do while not rs.eof and i < maxperpage 
				%>
              <TR height=20> 
                <TD align=middle width=90 bgColor=#ffffff height=26><B><FONT 
            face=Arial><%=rs("Admin_ID")%></FONT> </B></TD>
                <TD height="20" align=middle bgColor=#ffffff><%=rs("Admin_Username")%></TD>
                <TD height="20" align=middle bgColor=#ffffff><%=rs("Admin_TimeStart")%></TD>
                <TD height="20" align=middle bgColor=#ffffff><%=rs("Admin_TimeEnd")%></TD>
                <TD height="20" align=middle bgColor=#ffffff><%=rs("Admin_IP")%></TD>
                <TD align=middle bgColor=#ffffff> <input name="selCateID" type="checkbox" id="selCateID" value="<%=rs("Admin_ID")%>"> 
                </TD>
              </TR>
              <%
  i=i+1
    rs.movenext 
    loop  
   %>
              <TR> 
                <TD background=images/linedot.gif colSpan=6 
        height=1></TD>
              </TR>
            </TBODY>
          </TABLE>
          <table width="95%" border="0" cellpadding="0" cellspacing="1" >
            <td >共有记录：<font color="#FF0000"><%=totalput%></font>条&nbsp;页次：<font color="#FF0000"><%=CurrentPage%></font>/<font color="#FF0000"><%=TotalPages%></font> 
              页　 
              <select size="1" name="D1" onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
                <option selected>页数</option>
                <% 
		dim a
		dim b
        a=0
		b=TotalPages
      DO WHILE NOT a = b %>
                <option value="Admin_user.asp?key=<%=request("key")%>&Page=<% =a+1 %>"> 
                <% =a+1 %>
                </option>
                <% a=a+1
         loop    
      %>
              </select></td>
            <td align="right"> 
              <%
'------------------------------------------------------页面跳转
dim n,k 
	if (totalPut mod MaxPerPage)=0 then  'n表示总页数
		n= totalPut \ MaxPerPage
	else  
		n= totalPut \ MaxPerPage + 1  
	end if
k=currentPage
		if k<>1 then
			response.write "<a href=Admin_user.asp?key="&request("key")&"&page=1>首页</a>┊"
			response.write "<a href=Admin_user.asp?key="&request("key")&"&page="+cstr(k-1)+">上一页</a>┊"
		else
			Response.Write "<strong><font color=#999999>首页</font></strong>┊<strong><font color=#999999>上一页</font></strong>┊"
		end if
		if k<>n then
			response.write "<a href=Admin_user.asp?key="&request("key")&"&page="+cstr(k+1)+">下一页</a>┊"
			response.write "<a href=Admin_user.asp?key="&request("key")&"&page="+cstr(n)+">末页</a> "
		else
			Response.Write "<strong><font color=#999999>下一页</font></strong>┊<strong><font color=#999999>末页</font></strong> "
		end if
'------------------------------------------------------页面跳转
rs.close
set rs=nothing
%>
            </td>
            </tr>
          </table></FORM>
  </TD></TR></FORM></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
  <TR>
    <TD bgColor=#d6dff7 height=10></TD></TR></TBODY></TABLE>
<%	
end sub
sub add()
%>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="24" colspan="3" bgcolor="235ACB">&nbsp;<b><font color=#ffffff>管理员帐号管理</font></b></td>
  </tr>
  <tr bgcolor="#ffffff"> 
    <td width="5%">&nbsp;</td>
    <td width="90%" valign="top" bgcolor="#ffffff"> <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <form name=form onSubmit="return check()" method=post action=Admin_user.asp>
          <tr> 
            <td width="3%"><img src="images/meeting_memo.gif" width="16" height="16"></td>
            <td width="31%">添加新的管理员</td>
            <td width="66%">&nbsp;</td>
          </tr>
          <tr> 
            <td  background="images/linedot.gif" height=1 colspan="3"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>姓名： 
              <input name=Admin_Username class=lineinput id="Admin_Username"></td>
            <td rowspan="2" valign="bottom"><INPUT  name="submit" type=image src="images/yes.gif" align=absMiddle border=0  title=确定提交></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>密码： 
              <input name=Admin_Password type="password"  class=lineinput id="Admin_Password"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td> <input name=action type=hidden id="action" value="addUser"></td>
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
      </table></td>
    <td width="5%" valign="top" bgcolor="#ffffff">&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="3" height=10 bgcolor="#D6DFF7"></td>
  </tr>
</table>
<% end sub 
		sub Sysmsg(msgtitle,msginfo)
		%> <br>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TR> 
    <TD bgColor=#235acb height=24>&nbsp;<FONT 
    color=#ffffff><B>客户服务管理成功</B></FONT></TD>
  </TR>
</TABLE>
<TABLE cellSpacing=0 cellPadding=5 width="90%" align=center bgColor=#ffffff border=0>
  <TR> 
    <TD align=middle>&nbsp; </TD>
  </TR>
  <TR> 
    <TD><strong><%=msgtitle%></strong><br><%=msginfo%></TD>
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
	  <%end sub %> 
</BODY></HTML>
<%
CloseDatabase
%>