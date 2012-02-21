<!--#include file="conn/conn.asp"-->
<%
sql="select orders from Product_Info where ProductType_ID="&request("ProductType_ID")&" and Product_ID="&request("pid")
set rs=conn.execute(sql)
sorders=rs(0)

sql="select top 1 orders,Product_ID from Product_Info where ProductType_ID="&request("ProductType_ID")&" and orders>" & sorders & " order by orders"
set rs=conn.execute(sql)
if rs.eof then
	Response.Write "<script language=javascript>alert('该记录已移动到最上！');history.back()</script>"
	Response.End 
end if

corders=rs(0)
cid=rs(1)

sql="update Product_Info set orders="&corders&" where ProductType_ID="&request("ProductType_ID")&" and Product_ID="&request("pid")
conn.execute(sql)

sql="update Product_Info set orders="&sorders&" where ProductType_ID="&request("ProductType_ID")&" and Product_ID="&cid
conn.execute(sql)

Response.Redirect "Admin_product.asp?ProductType_ID="&request("ProductType_ID")&""
%>