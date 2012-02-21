<!-- #include file="conn/conn.asp" -->
<!-- #include file="inc/function.asp" -->
<!--#include file="inc/Char.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>华源皮具</title>
<link href="css/index.css" rel="stylesheet" type="text/css">
<script language=javascript>
<!--
   function checkwindow()
{
	if(window.screen.width==1024)
	{
		document.body.style.background="url('images/bk.jpg')";
	}
}  
 

//-->
</script>

</head>
<SCRIPT language=javascript>
 function Juge(){
    if (document.myform.Name.value==""){
       alert ("你的姓名不可为空！");
       document.myform.Name.focus();
       return(false);
    }
    if (document.myform.Tel.value==""){
       alert ("你的电话不可为空！");
       document.myform.Tel.focus();
       return(false);
    }
    if (document.myform.Inc.value==""){
       alert ("公司名称不可为空！");
       document.myform.Inc.focus();
       return(false);
    }
    if (document.myform.Email.value==""){
       alert ("Email不可为空！");
       document.myform.Email.focus();
       return(false);
    }
    if (document.myform.Add.value==""){
       alert ("地址不可为空！");
       document.myform.Add.focus();
       return(false);
    }
	
    if ((document.myform.Email.value.indexOf("@") == -1) || (document.myform.Email.value.indexOf(".") == -1)){
		alert("请查看您的E-mail地址是否正确，请重录入!");
		document.myform.Email.focus();
       return(false);
	}
	
	if (Check_Email(document.myform.Email.value)==true) {
		alert("请您正确填好电子邮件栏！");
		document.myform.Email.focus();		
       return(false);
	}	

    if (document.myform.Info.value==""){
       alert ("留言内容不可为空！");
       document.myform.Info.focus();
       return(false);
	}	
 }

function checkusername(text)
{
			allValid = false;
		if (text.length < 2)
		{
    		allValid = true;
		}

    	var notuser = "°′″＄￡￥‰％℃¤￠≈≡≠＝≤≥＜＞≮≯∷±＋－×÷／∫∮∝∞∧∨∑∏∪∩∈∵∴⊥∥∠⌒⊙≌∽√々＿￣〓＾＼→←↑↓※§№★☆○●◎◇◆□■△▲＃＆＠１２３４５６７８９０～！＂＇·＃￥％……ˇ＠¨〈〉「」『』．‖々〃〔〕〖〗—（），。【】《》？；‘：“”［］｛｝—＋＝｜｀、《》~`!@#$%^&*()_+|-=\'?/<>[],.:;";

		for (i = 0;  i < text.length;  i++)
		{
			for (j = 0;  j < notuser.length;  j++)
			{
              if (text.charAt(i) == notuser.charAt(j))
              {
				allValid = true;
				break;
              }
			}
			if (text.charAt(i) == " ")
			{
				allValid = true;
				break;
			}
		}

return allValid;
}

function teltext(text)
{
			allValid = false;
		if (text.length < 8)
		{
    		allValid = true;
		}
		if (text.length > 20)
		{
    		allValid = true;
		}
		
	var checkOK = "0123456789-_";

	for (i = 0; i < text.length;  i++)
	{
		ch = text.charAt(i);
		for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
		if (j == checkOK.length)
		{
  		allValid = true;
			break;
		}
	}
return allValid;
}

function checktext(text)
{
			allValid = false;
		if (text.length < 2)
		{
    		allValid = true;
		}
		if (text.length > 12)
		{
    		allValid = true;
		}
		
	var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";

	for (i = 0;  i < text.length;  i++)
	{
		ch = text.charAt(i);
		for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
		if (j == checkOK.length)
		{
  		allValid = true;
			break;
		}
	}
return allValid;
}

   function Check_Email(string){ 
	  var str_len = string.length;
	  if (str_len<=5){
       return(false);
    	}  
	  for(i=0;i<str_len;i++){
	     if (string.charCodeAt(i)>127){
       return(false);
		}
	  }
	  if (string.indexOf("@")<2){
       return(false);
    	}  
        if (string.indexOf(".")<4){
       return(false);
    	}  
	  if (string.indexOf(":")!=-1){
       return(false);
	  }
   }
</SCRIPT>
<body  onload="checkwindow()">
<table width="780" height="97" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td><img src="images/bk_04.jpg" width="780" height="97"></td>
  </tr>
</table>
<table width="780" height="51" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="209"><img src="images/bk_09.jpg" width="209" height="51"></td>
    <td><img src="images/bk_13.jpg" width="531" height="51"></td>
    <td><img src="images/bk_15.jpg" width="40" height="51"></td>
  </tr>
</table>
<table width="780" height="342" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="209">
<table width="209" height="342" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="114" align="left" valign="top"><img src="images/bk_10.jpg" width="209" height="114"></td>
        </tr>
        <tr>
          <td align="left"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="209" height="228">
              <param name="movie" value="flash/menu.swf">
              <param name="quality" value="high">
              <embed src="flash/menu.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="209" height="228"></embed></object></td>
        </tr>
      </table>
    </td>
    <td><table width="531" height="342" border="0" cellpadding="0" cellspacing="0" background="images/bk_20.jpg">
        <tr>
          <td height="60">&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td width="70">&nbsp;</td>
          <td>
            <%	
  if request("action")="add" then
       call SaveAdd()
  'elseif request("action")="post" then
  '	   call mypost()
  else
  		call mypost()
       'call myform()
  end if
  
	
	sub SaveAdd()
      set rs=server.createobject("adodb.recordset")
	  sql="select * from GuestBook where (GuestBook_ID is null)"
	  rs.open sql,conn,1,3
  	   rs.addnew
	   rs("Name")=checkstr(Trim(request.form("Name")))
	   rs("Tel")=trim(request.form("Tel"))
	   rs("Inc")=trim(request.form("Inc"))
	   rs("Email")=Trim(Request.Form("Email"))
	   rs("Add")=trim(request.form("Add"))
	   rs("Info")=request.form("Info")
	   rs("Time")=Now()
	   rs.update
	   call Sysmsg()
	end sub
  
sub mypost()
%>
            <table width="362" height="122" border="0" align="center" cellpadding="0" cellspacing="0" class="text">
                <form action="feedback.asp" method="post" name="myform" id="myform" onSubmit="return Juge(this)"><tr> 
                  <td width="70"> 姓　　名:</td>
                  <td width="292" height="20"><input name="Name" type="text" id="Name2">
                    　* </td>
                </tr>
                <tr> 
                  <td height="1" colspan="2"><img src="images/line.gif" width="362" height="1"></td>
                </tr>
                <tr> 
                  <td>电　　话:</td>
                  <td height="20"><input name="Tel" type="text" id="Tel2">
                    　* </td>
                </tr>
                <tr> 
                  <td height="1" colspan="2"><img src="images/line.gif" width="362" height="1"></td>
                </tr>
                <tr> 
                  <td>公司名称:</td>
                  <td height="20"><input name="Inc" type="text" id="Inc2">
                    　* </td>
                </tr>
                <tr> 
                  <td height="1" colspan="2"><img src="images/line.gif" width="362" height="1"></td>
                </tr>
                <tr> 
                  <td>Email 　:</td>
                  <td height="20"><input name="Email" type="text" id="Email2">
                    　* </td>
                </tr>
                <tr> 
                  <td height="1" colspan="2"><img src="images/line.gif" width="362" height="1"></td>
                </tr>
                <tr> 
                  <td>址　　址:</td>
                  <td height="20"><input name="Add" type="text" id="Add2">
                    　* </td>
                </tr>
                <tr> 
                  <td height="1" colspan="2"><img src="images/line.gif" width="362" height="1"></td>
                </tr>
                <tr> 
                  <td valign="middle">留言内容:</td>
                  <td height="110"><textarea name="Info" cols="30" rows="6" id="textarea"></textarea>
                    <input name="action" type="hidden" id="action7" value="add" /></td>
                </tr>
                <tr align="center"> 
                  <td height="30" colspan="2"><input name="imageField" type="image" src="images/bt_01.jpg" width="70" height="31" border="0"></td>
                </tr></form>
              </table>
            <%
				  end sub
				  sub Sysmsg()
				  %>
            <table width="362" height="150" border="0" align="center" cellpadding="5" cellspacing="0" class="text">
              <tr> 
                <td>恭喜！！！<br />
                  留言发表成功！！！<br />
                  请在下面列表查看你发表的留言！！！<br />
                  你的留言在我们查看后会第一时间回复！！！<br /> <br />
                  [<a href="feedback.asp">我要继续留言</a>] </td>
              </tr>
            </table>
            <%end sub%>
          </td>
          <td width="60" class="text">&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td height="30">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td width="40">
<table width="40" height="342" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="114"><img src="images/bk_16.jpg" width="40" height="114"></td>
        </tr>
        <tr>
          <td><img src="images/bk_17.jpg" width="40" height="228"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="780" height="65" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="209"><img src="images/bk_12.jpg" width="209" height="65"></td>
    <td><img src="images/bk_14.jpg" width="531" height="65" border="0" usemap="#Map"></td>
    <td><img src="images/bk_18.jpg" width="40" height="65"></td>
  </tr>
</table>
<map name="Map">
  <area shape="rect" coords="261,43,522,61" href="http://www.vtrue.net" target="_blank">
</map>
</body>
</html>
