<!-- #include file="conn/conn.asp" -->
<!--#include file="../inc/function.asp" -->
<!--#include file="../inc/Char.asp" -->
<HTML><HEAD>
<STYLE type=text/css>BODY {
	BACKGROUND: #799ae1; MARGIN: 0px; FONT: 9pt ����
}
TABLE {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
TD {
	FONT: 12px Tahoma,����
}
IMG {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; VERTICAL-ALIGN: bottom; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
A {
	FONT: 12px Tahoma,����; COLOR: #000000; TEXT-DECORATION: none
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

<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"block\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
function op(k)
{
   for (i=0;i<=k;i++){
   whichEl = eval("submenu" + i)
   if (whichEl.style.display=="none") whichEl.style.display="block"}
   }
function cl(k)
{
   for (i=0;i<=k;i++){
   whichEl = eval("submenu" + i)
   if (whichEl.style.display=="block") whichEl.style.display="none"}
   }

function oa_tool(){
if(window.parent.oa_frame.cols=="2,18,*"){
//frameshow.src="up.gif";
oa_tree.title="��ʾ������"
window.parent.oa_frame.cols="175,18,*";
}
else{
//frameshow.src="up.gif";
oa_tree.title="���ع�����"
window.parent.oa_frame.cols="0,18,*";}
}
</SCRIPT>

<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2800.1479" name=GENERATOR><title>����ҳ��</title></HEAD>
<BODY leftMargin=0 topMargin=0 marginwidth="0" marginheight="0">
<DIV align=center>
<CENTER>
<TABLE 
style="BORDER-TOP-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; BORDER-COLLAPSE: collapse; BORDER-RIGHT-WIDTH: 0px" 
borderColor=#111111 height="100%" cellSpacing=0 cellPadding=0>
  <TBODY>
  <TR>
    <TD vAlign=top>
      <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#111111 height="100%" 
      cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD vAlign=top>
            <DIV align=center>
            <CENTER>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#111111 
            cellSpacing=0 cellPadding=0 width=157>
              <TBODY>
              <TR>
                              <TD width=159 
                height=38 vAlign=bottom background=images/title.gif> 
                                <DIV style="WIDTH: 156px; HEIGHT: 20px" align=center><strong><font color="#FFFFFF">��̨����&nbsp;&nbsp;</font></strong></DIV></TD></TR></TBODY></TABLE></CENTER></DIV>
            <DIV align=center>
            <CENTER>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#111111 
            cellSpacing=0 cellPadding=0 width=157>
              <TBODY>
              <TR>
                              <TD class=menu_title onmouseover="this.className='menu_title2';" 
                onmouseout="this.className='menu_title';" 
                background=images/title_bg_quit.gif 
                  height=30><SPAN>&nbsp;<A onclick=op(9) href="left.asp#"><B>��ȫ��</B></A> 
                                &nbsp;&nbsp;<A onclick=cl(9) href="left.asp#"><B>�ر�ȫ��</B></A></SPAN></TD>
                            </TR></TBODY></TABLE></CENTER></DIV>
                    <TABLE cellSpacing=0 cellPadding=0 width=157 align=center>
                      <TBODY>
              <TR>
                          <TD class=menu_title id=menuTitle1 
                onmouseover="this.className='menu_title2';" 
                onclick=showsubmenu(0) onmouseout="this.className='menu_title';" 
                background=images/admin_left_7.gif 
                  height=25><SPAN>��Ʒ��������</SPAN> </TD>
                        </TR>
              <TR>
                <TD id=submenu0 style="DISPLAY: none">
                  <DIV class=sec_menu style="WIDTH: 157px">
                              <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                                <TBODY>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from Product_type",conn,1,1
		  dim follows
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>��û�з���</font></center>"
		  else
		  do while not rs.EOF
		  %>
                                  <TR> 
                                    <TD height=20><A href="Admin_product.asp?ProductType_ID=<%=rs("ProductType_ID")%>" target=BoardList><%=rs("Product_type")%></A></TD>
                                  </TR>
        <%rs.MoveNext
          loop
          end if
		  %>
                                  <TR>
                                    <TD height=20><A 
                        href="Admin_addproduct.asp" 
                        target=BoardList>��Ӳ�Ʒ</A></TD>
                                  </TR>
                                  <TR> 
                                    <!-- <TD height=20><A 
                        href="Admin_producttype.asp" 
                        target=BoardList>��Ʒ���</A></TD> -->
                                  </TR>
                                </TBODY>
                              </TABLE>
                            </DIV>
                  <DIV style="WIDTH: 157px">
                  <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                    <TBODY>
                    <TR>
                      <TD 
            height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
<!--                     <TABLE cellSpacing=0 cellPadding=0 width=157 align=center>
              <TBODY>
              <TR>
                          <TD class=menu_title id=menuTitle1 
                onmouseover="this.className='menu_title2';" 
                onclick=showsubmenu(1) onmouseout="this.className='menu_title';" 
                background=images/admin_left_2.gif 
                  height=25><SPAN>���Ź���</SPAN> </TD>
                        </TR>
              <TR>
                <TD id=submenu1 style="DISPLAY: none">
                  <DIV class=sec_menu style="WIDTH: 157px">
                              <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                                <TBODY>
                                  <TR> 
                                    <TD height=20><a href="Admin_Infolist.asp?Type_ID=1" target="BoardList">��˾���Ź���</a></TD>
                                  </TR>
                                  <TR> 
                                    <TD height=20><a href="Admin_info.asp?Type_ID=1" target="BoardList">��˾�������</a></TD>
                                  </TR>
                                  <TR> 
                                    <TD height=20><a href="Admin_Infolist.asp?Type_ID=2" target="BoardList">��Ʒ���й���</a></TD>
                                  </TR>
                                  <TR> 
                                    <TD height=20><a href="Admin_info.asp?Type_ID=2" target="BoardList">��Ʒ�������</a></TD>
                                  </TR>
                                  <TR> 
                                    <TD height=20><a href="Admin_Infolist.asp?Type_ID=3" target="BoardList">��ҵ��̬����</a></TD>
                                  </TR>
                                  <TR> 
                                    <TD height=20><a href="Admin_info.asp?Type_ID=3" target="BoardList">��ҵ��̬���</a></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
                            </DIV>
                  <DIV style="WIDTH: 157px">
                  <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                    <TBODY>
                    <TR>
                      <TD 
            height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=157 align=center>
              <TBODY>
              <TR>
                          <TD class=menu_title id=menuTitle2
                onmouseover="this.className='menu_title2';" 
                onclick=showsubmenu(2) onmouseout="this.className='menu_title';" 
                background=images/admin_left_3.gif 
                  height=25><SPAN>���߶�������</SPAN> </TD>
                        </TR>
              <TR>
                <TD id=submenu2 style="DISPLAY: none">
                  <DIV class=sec_menu style="WIDTH: 157px">
                              <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                                <TBODY>
                                  <TR> 
                                    <TD height=20><a href="Admin_Order.asp" target="BoardList">���߶�������</a></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
                            </DIV>
                  <DIV style="WIDTH: 157px">
                  <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                    <TBODY>
                    <TR>
                      <TD 
            height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE> -->            
                    <TABLE cellSpacing=0 cellPadding=0 width=157 align=center>
                      <TBODY>
              <TR>
                          <TD class=menu_title id=menuTitle3
                onmouseover="this.className='menu_title2';" 
                onclick=showsubmenu(3) onmouseout="this.className='menu_title';" 
                background=images/admin_left_6.gif 
                  height=25><SPAN>���Թ���</SPAN> </TD>
                        </TR>
              <TR>
                <TD id=submenu3 style="DISPLAY: none">
                  <DIV class=sec_menu style="WIDTH: 157px">
                              <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                                <TBODY>
                                  <TR> 
                                    <TD height=20><A 
                        href="Admin_guestbook.asp" target="BoardList">������Ϣ����</A></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
                            </DIV>
                  <DIV style="WIDTH: 157px">
                  <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                    <TBODY>
                    <TR>
                      <TD 
            height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
                    <!-- <TABLE cellSpacing=0 cellPadding=0 width=157 align=center>
                      <TBODY>
              <TR>
                          <TD class=menu_title id=menuTitle4
                onmouseover="this.className='menu_title2';" 
                onclick=showsubmenu(4) onmouseout="this.className='menu_title';" 
                background=images/admin_left_8.gif 
                  height=25><SPAN>��������</SPAN> </TD>
                        </TR>
              <TR>
                <TD id=submenu4 style="DISPLAY: none">
                  <DIV class=sec_menu style="WIDTH: 157px">
                              <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                                <TBODY>
                                  <TR> 
                                    <TD height=20><A 
                        href="Admin_commend.asp?Info_ID=1" 
                        target=BoardList>��վTitle����</A></TD>
                                  </TR>
                                </TBODY>
                              </TABLE>
                            </DIV>
                  <DIV style="WIDTH: 157px">
                  <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                    <TBODY>
                    <TR>
                      <TD 
            height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE> -->
			<TABLE cellSpacing=0 cellPadding=0 width=157 align=center>
              <TBODY>
              <TR>
                          <TD class=menu_title id=menuTitle5
                onmouseover="this.className='menu_title2';" 
                onclick=showsubmenu(5) onmouseout="this.className='menu_title';" 
                background=images/admin_left_10.gif 
                  height=25><SPAN>��վϵͳ����</SPAN> </TD>
                        </TR>
              <TR>
                <TD id=submenu5 style="DISPLAY: none">
                  <DIV class=sec_menu style="WIDTH: 157px">
                  <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                    <TBODY>
                    <TR>
                                    <TD height=20><A 
                        href="Admin_user.asp" 
                        target=BoardList>�༭����Ա�ʺ�</A></TD>
                                  </TR>
                    <TR>
                                    <TD height=20><A 
                        href="Admin_useredit.asp" 
                        target=BoardList>�޸Ĺ���Ա�ʺ�</A></TD>
                                  </TR>
                    <TR>
                                    <TD height=20><A 
                        href="login.asp?action=logout" 
                        target=_parent>�˳�����ϵͳ</A></TD>
                                  </TR></TBODY></TABLE></DIV>
                  <DIV style="WIDTH: 157px">
                  <TABLE cellSpacing=0 cellPadding=0 width=135 align=center>
                    <TBODY>
                    <TR>
                      <TD 
            height=20></TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
            <DIV align=center>
            <CENTER>
            <TABLE style="BORDER-COLLAPSE: collapse" borderColor=#111111 
            cellSpacing=0 cellPadding=0 width=157>
              <TBODY>
              <TR>
                <TD class=menu_title id=menuTitle6
                onmouseover="this.className='menu_title2';" 
                onmouseout="this.className='menu_title';" 
                background=images/title_bg_quit.gif>��</TD></TR>
              <TR>
                <TD>
                  <DIV class=sec_menu style="WIDTH: 157px">
                  <TABLE cellSpacing=0 cellPadding=0 width=155 align=center>
                    <TBODY>
                    <TR>
                                        <TD align=center width=153 height=30>�����л�ԴƤ�����޹�˾</TD>
                                      </TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE></CENTER></DIV></TD>
          <TD height="100%"></TD></TR></TBODY></TABLE>
      <P>&nbsp;</P>
      <DIV></DIV></TR></TBODY></TABLE></CENTER></DIV></BODY></HTML>
