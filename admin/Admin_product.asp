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
  case "del"
       call delCate()
  case else
       call myform() 
  end select
else
		response.Redirect"login.asp"		   	   
end if
  
  sub delCate()
        conn.execute("delete from Product_Info where Product_ID in ("&Request.Form("selCateID")&")")
        call Sysmsg() 
  end sub
  %> <% sub myform() %> 
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
  <TR>
      <TD bgColor=#235acb height=24>&nbsp;<FONT color=#ffffff><B>产品管理</B></FONT></TD>
    </TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
  <TR>
      <TD align=middle bgColor=#ffffff><form name="selform" method="post" action="Admin_product.asp" >
          <TABLE cellSpacing=0 cellPadding=3 width="98%" border=0>
            <TBODY>
              <TR> 
                <TD width="5%"><IMG src="images/search.gif"></TD>
                <TD width="95%">&nbsp;关键字: 
                  <INPUT name=Key class=lineInput id="Key"> &nbsp;&nbsp; 按名称查询&nbsp; 
                  <INPUT type=image height=18 width=43 
            src="images/search1.gif" align=absMiddle border=0 name=image onClick="return cha()"> 
                </TD>
              </TR>
            </TBODY>
          </TABLE>
        </form> 
        <form name="form" method="post" action="Admin_product.asp">
          <TABLE cellSpacing=0 cellPadding=4 width="95%" border=0>
            <TBODY>
              <TR> 
                <TD background=images/linedot.gif colSpan=7 height=1></TD>
              </TR>
              <TR height=25> 
                <TD width="90" height="25" align=left><FONT color=#ff6600><B>ID号</B></FONT></TD>
                <TD height="25"><FONT color=#ff6600><B>产品名称(英文名称)</B></FONT></TD>
                <TD height="25"><FONT color=#ff6600><B>查看</B></FONT></TD>
                <TD width="80" height="25" align=middle><FONT color=#ff6600><B>日期</B></FONT></TD>
                <TD width="50" align=center><font color="#ff6600"><strong>排序</strong></font></TD>
                <TD width="30" height="25" align=middle><FONT color=#ff6600><B>修改</B></FONT></TD>
                <TD width=40 height="25" align=middle><input name=action type=hidden id="action2" value="del"> 
                  <INPUT name="submit" type=submit class=buttonface value=删除></TD>
              </TR>
              <TR> 
                <TD background=images/linedot.gif colSpan=7 height=1></TD>
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
				sql = "select * from Product_Info where ProductType_ID="&request("ProductType_ID")&" order by orders Desc"
				else
				sql = "select * from Product_Info where ProductType_ID="&request("ProductType_ID")&" and Product_Name like '%"&Trim(Request("Key"))&"%' order by orders Desc"
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
              <TR> 
                <TD width=43 height=20 align=middle bgColor=#ffffff><B><FONT 
            face=Arial><%=rs("Product_ID")%></FONT> </B></TD>
                <TD bgColor=#ffffff> <%=rs("Product_Name")%></TD>
                <TD bgColor=#ffffff><a href="Admin_productshow.asp?Product_ID=<%=rs("Product_ID")%>"><img src="images/search.gif" width="28" height="22" border="0"></a></TD>
                <TD align=middle bgColor=#ffffff><%=DateTimeFormat(rs("Product_Time"),3)%></TD>
                <TD align=center bgColor=#ffffff><a href=jsup.asp?ProductType_ID=<%=rs("ProductType_ID")%>&pid=<%=rs("Product_ID")%>>↓</a>&nbsp;&nbsp;<a href=jsdown.asp?ProductType_ID=<%=rs("ProductType_ID")%>&pid=<%=rs("Product_ID")%>>↑</a></TD>
                <TD width="26" align=middle bgColor=#ffffff><A 
            href='Admin_addproduct.asp?action=edit&Product_ID=<%=rs("Product_ID")%>'><IMG height=14 
            src="images/m_news.gif" width=21 border=0></A></TD>
                <TD align=middle bgColor=#ffffff> <input name="selCateID" type="checkbox" id="selCateID2" value="<%=rs("Product_ID")%>"> 
                </TD>
              </TR>
              <TR> 
                <TD background=images/linedot.gif colSpan=7 height=1></TD>
              </TR>
              <%
  i=i+1
    rs.movenext 
    loop  
   %>
            </TBODY>
          </TABLE>
        </form>
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
                <option value="Admin_product.asp?ProductType_ID=<%=request("ProductType_ID")%>&key=<%=request("key")%>&Page=<% =a+1 %>"> 
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
			response.write "<a href=Admin_product.asp?ProductType_ID="&request("ProductType_ID")&"&key="&request("key")&"&page=1>首页</a>┊"
			response.write "<a href=Admin_product.asp?ProductType_ID="&request("ProductType_ID")&"&key="&request("key")&"&page="+cstr(k-1)+">上一页</a>┊"
		else
			Response.Write "<strong><font color=#999999>首页</font></strong>┊<strong><font color=#999999>上一页</font></strong>┊"
		end if
		if k<>n then
			response.write "<a href=Admin_product.asp?ProductType_ID="&request("ProductType_ID")&"&key="&request("key")&"&page="+cstr(k+1)+">下一页</a>┊"
			response.write "<a href=Admin_product.asp?ProductType_ID="&request("ProductType_ID")&"&key="&request("key")&"&page="+cstr(n)+">末页</a> "
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
  </TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
  <TR>
    <TD bgColor=#d6dff7 height=10></TD></TR></TBODY></TABLE>
<% end sub 
		sub Sysmsg()
		%> <br>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TR> 
    <TD bgColor=#235acb height=24>&nbsp;<FONT 
    color=#ffffff><B>删除成功</B></FONT></TD>
  </TR>
</TABLE>
<TABLE cellSpacing=0 cellPadding=5 width="90%" align=center bgColor=#ffffff border=0>
  <TR> 
    <TD >&nbsp; </TD>
  </TR>
  <TR> 
    <TD><li>产品删除成功！！！</li></TD>
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