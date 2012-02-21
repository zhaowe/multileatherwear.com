<%@LANGUAGE="VBSCRIPT"%>
<%
	'option explicit
	dim conn
	dim connstr
	dim db
	'更改数据库名字
	db="huayuan&data/huayuan&data.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
	conn.Open connstr

function CloseDatabase

	Conn.close
	Set conn = Nothing

End Function
Dusername=session("username")
%>