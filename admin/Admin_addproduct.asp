<!-- #include file="conn/conn.asp" -->
<!--#include file="../inc/function.asp" -->
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="style.css" type=text/css rel=stylesheet>
<script language="Javascript">
function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}
function openem()
{ 
openScript('uploadCP.asp',350,200); 
}
</script>
<SCRIPT language=JavaScript>
function Juge(myform)
{
  //if (document.myform.ProductType_ID.value=="0")
 //   {
  //   alert("请选择类别！")
     //document.myform.Product_Name.focus()
  //   document.myform.Product_Name.select()
  //   return(false);
  //   }

  if (document.myform.Product_Number.value=="")
    {
     alert("请输入产品名称！")
     document.myform.Product_Name.focus()
     document.myform.Product_Name.select()
     return(false);
     }

  if (document.myform.Product_Smaill.value=="")
    {
     alert("你还没有上传产品缩略图呢！！！")
     document.myform.Product_Smaill.focus()
     document.myform.Product_Smaill.select()		
     return(false);
    }
  if (document.myform.Product_Big.value=="")
    {
     alert("你还没有上传产品的大图呢！！！")
     document.myform.Product_Big.focus()
     document.myform.Product_Big.select()		
     return(false);
    }
  if (eWebEditor1.getHTML()=="")
    {
     alert("你还没有填写内容呢！！！")
     return (false);
    }
		eWebEditor1.remoteUpload("doSubmit()");
		return false;
}
	function doSubmit(){
		document.myform.submit();
	}
</SCRIPT>

<SCRIPT language=javascript>
  var upfile_obj;
  function OnUpFile()
  {
   upfile_obj.src=upfile_obj.lowsrc;
   upfile_obj.src=document.forms["myform"].upfile.value;
  }
</SCRIPT>

<STYLE type=text/css>.style1 {
	COLOR: #ff0000
}
</STYLE>

<META content="MSHTML 6.00.2900.2523" name=GENERATOR></HEAD>
<BODY bgColor=#799ae1 leftMargin=0 marginwidth="0" marginheight="0" onkeydown='if(event.keyCode==13 && event.ctrlKey)myform.submit()'>
<%	
  if not isnull(Admin_Username) and Admin_Username<>"" then 
  select case request("action")
  case "add" 
       call SaveAdd()
  case "modify"  
       call SaveModify()
  case "edit" 
       isEdit=True
       call myform(isEdit)
  case else
       isEdit=False
       call myform(isEdit)
  end select 
  else  	  
		response.Redirect"login.asp"
  end if
        
  sub SaveModify()
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next
      set rs=server.createobject("adodb.recordset")
	  sql="select * from Product_Info where Product_ID="&request.form("Product_ID")
	   rs.open sql,conn,1,3
	   rs("ProductType_ID")=trim(request.form("ProductType_ID"))
	   rs("Product_Name")=request.form("Product_Name")
	   rs("Product_Color")=trim(request.form("Product_Color"))
	   rs("Product_Size")=trim(request.form("Product_Size"))
	   rs("Product_Smaill")=trim(request.form("Product_Smaill"))
	   rs("Product_Big")=trim(request.form("Product_Big"))
	   rs("Product_Info")=sContent
	   rs("Product_Time")=Now()
	   rs.update
	   rs.close
	   set rs=nothing
	   msgtitle="保存修改成功"
	   msginfo="<li>产品修改成功！！！</li><li><a href=""Admin_Product.asp?ProductType_ID="&request("ProductType_ID")&""">继续修改产品</a></li>"
	   call Sysmsg(msgtitle,msginfo) 
	end sub 
	
	sub SaveAdd()
	For i = 1 To Request.Form("d_content").Count
		sContent = sContent & Request.Form("d_content")(i)
	Next
      set rs=server.createobject("adodb.recordset")
	  sql="select * from Product_Info where (Product_ID is null)"
	  rs.open sql,conn,1,3
  	   rs.addnew
	   rs("ProductType_ID")=trim(request.form("ProductType_ID"))
	   rs("Product_Name")=request.form("Product_Name")
	   rs("Product_Color")=trim(request.form("Product_Color"))
	   rs("Product_Size")=trim(request.form("Product_Size"))
	   rs("Product_Smaill")=trim(request.form("Product_Smaill"))
	   rs("Product_Big")=trim(request.form("Product_Big"))
	   rs("Product_Info")=sContent
	   rs("Product_Time")=Now()
	   rs.update
		  set rs= server.createobject("adodb.recordset")
		  sql="select top 1 * from Product_Info order by Product_ID desc"
		  rs.Open sql,conn,2,3
		  tempid=rs(0)
		  rs("orders").value=tempid
		  rs.Update
	   
	   rs.close
	   set rs=nothing
	   msgtitle="添加产品成功"
	   msginfo="<li>产品添加成功！！！</li><li><a href=""Admin_addproduct.asp?ProductType_ID="&request("ProductType_ID")&""">继续添加产品</a></li><li><a href=""Admin_product.asp"">返回产品列表页</a></li>"
	   call Sysmsg(msgtitle,msginfo)
	end sub
  
sub myform(isEdit)

%>
<form name="myform" method="post" action="Admin_addproduct.asp" onSubmit="return Juge(this)">
  <%
  		 if isedit then
		  set rse=server.createobject("adodb.recordset")
 		   rse.open "select * from Product_Info where Product_ID="&cstr(request("Product_ID")),conn,1,1
	        pubTitle="编辑产品"
			ProductType_ID=rse("ProductType_ID")
	     else
	       pubTitle="添加产品"
			ProductType_ID=0
	    end if %>
<input type="Hidden" name="action" value='<% If isedit then%>modify<% Else  %>add<% End If %>'>
<%If isedit then%>
<input type="Hidden" name="Product_ID" value='<%=cstr(request("Product_ID"))%>'>
<%End If%>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
    <TR> 
      <TD bgColor=#235acb height=24>&nbsp;<FONT 
    color=#ffffff><B><%=pubTitle%></B></FONT></TD>
    </TR>
  </TBODY>
</TABLE><TABLE cellSpacing=0 cellPadding=0 width="90%" align=center bgColor=#ffffff 
border=0>
<TBODY>
  <TR> 
    <TD align=middle> <TABLE cellSpacing=0 cellPadding=2 width="100%" border=0>
            <TBODY>
              <TR> 
                <TD align=right width="120">&nbsp;</TD>
                <TD width="32%">&nbsp;</TD>
                <TD width="13%">&nbsp;</TD>
                <TD width="40%">&nbsp;</TD>
              </TR>
              <TR> 
                <TD align=right>产品类别：</TD>
                <TD colspan="3"> <select name="ProductType_ID" size="1">
                    <%
			    set rs=server.createobject("adodb.recordset")
				sql="select * from Product_type order by orders Asc"
				rs.open sql,conn,1,1
				if rs.eof and rs.bof then
					response.write "没有添加类别"
				else
				do while not rs.eof
					if cstr(ProductType_ID)=cstr(rs("ProductType_ID")) then 
					sel="selected"
					else
					sel=""
					end if	
				response.write "<option " & sel & " value='"&cstr(rs("ProductType_ID"))&"'>"+trim(rs("Product_Type"))+"</option>"
				rs.movenext
					loop
				end if
				rs.close
				%>
                  </select> </TD>
              </TR>
              <TR> 
                <TD align=right><font color="#FF0000">* </font>产品名称：</TD>
                <TD colspan="3"><input name=Product_Name id="Product_Number" value="<% if isedit then response.write trim(rse("Product_Name")) end if %>" size=28></TD>
              </TR>
              <TR> 
                <TD align=right>颜 色：</TD>
                <TD colspan="3"><INPUT name=Product_Color id="Product_Color" value="<% if isedit then response.write trim(rse("Product_Color")) end if %>" size=28></TD>
              </TR>
              <TR> 
                <TD align=right>尺 码：</TD>
                <TD colspan="3"><INPUT name=Product_Size id="Product_Size" value="<% if isedit then response.write trim(rse("Product_Size")) end if %>" size=28></TD>
              </TR>
              <TR> 
                <TD align=right><font color="#FF0000">* </font>产品小图：</TD>
                <TD colspan="3"><INPUT name=Product_Smaill id="Product_Smaill" value="<% if isedit then response.write trim(rse("Product_Smaill")) else response.Write("none.gif") end if %>" size=28> 
                  <font color="#FF0000">大小最好宽为100px高不定</font> [<a href="JavaScript:openScript('uploadCP.asp',350,200)">上传图片</a>]</TD>
              </TR>
              <TR> 
                <TD align=right><font color="#FF0000">* </font>产品大图：</TD>
                <TD colspan="3"><INPUT name=Product_Big id="Product_Big" value="<% if isedit then response.write trim(rse("Product_Big")) else response.Write("none.gif") end if %>" size=28> 
                  <font color="#FF0000">大小最好宽为500px高不定</font></TD>
              </TR>
              <TR> 
                <TD></TD>
                <TD colSpan=3></TD>
              </TR>
              <TR> 
                <TD align=right valign="top">产品简介：</TD>
                <TD colSpan=3><textarea name="d_content" style="display:none"><% if isedit then response.write Server.HtmlEncode(rse("Product_Info")) end if %></textarea> 
                  <iframe ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=d_content&style=s_coolblue" frameborder="0" scrolling="no" width="550" HEIGHT="350"></iframe></TD>
              </TR>
              <TR> 
                <TD align=right>&nbsp;</TD>
                <TD align=right>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
              </TR>
              <TR> 
                <TD align=right>&nbsp;</TD>
                <TD align=right><INPUT name=enter type=submit class=buttonface id="enter" value=确定> 
                  &nbsp;&nbsp; </TD>
                <TD>&nbsp; <INPUT class=buttonface type=reset value="清除" name=Submit2> 
                </TD>
                <TD>&nbsp;</TD>
              </TR>
              <TR> 
                <TD colSpan=4></TD>
              </TR>
          </TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TBODY>
  <TR>
    <TD bgColor=#d6dff7 height=10></TD></TR></TBODY></TABLE> </FORM>
        <% 
			'rse.Close
			set rse=nothing
			end sub 
		sub Sysmsg(msgtitle,msginfo)
		%>
<BR>
<TABLE cellSpacing=0 cellPadding=0 width="90%" align=center border=0>
  <TR> 
    <TD bgColor=#235acb height=24>&nbsp;<FONT 
    color=#ffffff><B><%=msgtitle%></B></FONT></TD>
  </TR>
</TABLE>
<TABLE cellSpacing=0 cellPadding=5 width="90%" align=center bgColor=#ffffff border=0>
  <TR> 
    <TD align=middle>&nbsp; </TD>
  </TR>
  <TR> 
    <TD><%=msginfo%></TD>
  </TR>
  <TR> 
    <TD align=middle>&nbsp;</TD>
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
<br> <%end sub %> 
</BODY></HTML>
<%
CloseDatabase
%>