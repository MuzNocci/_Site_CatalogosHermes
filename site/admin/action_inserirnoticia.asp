<%
noticia = Request.form("noticia")
titulo = Request.form("titulo")

Set Con=Server.CreateObject("ADODB.Connection")
Con.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("../../../ch_database/bd_hermes.mdb")

SqlString = "INSERT INTO noticias (titulo,noticia,data,hora) VALUES ('"&titulo&"','"&noticia&"','"&Date()&"', '"&Time()&"')"
Con.Execute sqlString
Con.Close
if err = 0 Then
	response.redirect "sucesso.asp"
end if
%> 