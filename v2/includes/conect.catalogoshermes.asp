<%
IF Application("fwsseg") <> "fwshermeshagf87agsfqgwrgv67g" THEN

response.redirect("/site/")

ELSE

Set objConn =  Server.CreateObject("ADODB.Connection")
objConn.Open "DBQ=" & Server.MapPath("../../ch_database/bd_hermes.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","Admin",""

END IF
%>