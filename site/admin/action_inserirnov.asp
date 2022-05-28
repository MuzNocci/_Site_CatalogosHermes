<%
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open "DBQ=" & Server.MapPath("../../../ch_database/bd_hermes.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","admin",""

Application("titulo")   = trim(Upload.form("titulo"))
Application("novidade") = trim(Upload.form("novidade"))

Set Upload = Server.CreateObject("Persits.Upload")
Upload.Save(""&Server.MapPath("img/")&"")

Set File1 = Upload.files("FILE1")
if not File1 is nothing then
Application("file1") = File1.Filename
end if

inserir = "INSERT INTO novidades (titulo,novidade,imagem) VALUES ('"&Application("titulo")&"','"&Application("novidade")&"','"&Application("file1")&"')"
Conexao.Execute(inserir)

Set File1 = Nothing
Conexao.Close
Set Conexao = Nothing

Response.Redirect "sucesso.asp"
%>