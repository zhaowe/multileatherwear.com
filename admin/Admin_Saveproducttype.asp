<!-- #include file="conn/conn.asp" -->
<%
dim action,categoryid
categoryid=Request("id")
action=Request("action")
select case action

case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from Product_type",conn,1,3
rs.AddNew
rs("Product_type")=trim(request.form("Product_type"))
rs("Orders")=trim(request.form("Orders"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "Admin_producttype.asp"

case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from Product_type where ProductType_ID="&categoryid,conn,1,3
rs("Product_type")=trim(request.form("Product_type"))
rs("Orders")=trim(request.form("Orders"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "Admin_producttype.asp"

case "del"
conn.execute ("delete from Product_type where ProductType_ID="&categoryid)
conn.execute ("delete from Product_Info where ProductType_ID="&categoryid)
response.Redirect "Admin_producttype.asp"
end select
%>