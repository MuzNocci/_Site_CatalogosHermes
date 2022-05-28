<%@LANGUAGE="VBSCRIPT"%>
<% Server.ScriptTimeout = 300 %>
<%
Dim objConn, stringSQL, strConnection, array_id, i, sql_id, id
id = cdbl(Request.QueryString("radio"))
Set objConn =  Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("../../../ch_database/bd_hermes.mdb")
stringSQL = "DELETE * FROM datas WHERE id ="&id
objConn.Execute(stringSQL)		
objConn.close
Set objConn = Nothing		  
response.Redirect "sucesso.asp"
%>