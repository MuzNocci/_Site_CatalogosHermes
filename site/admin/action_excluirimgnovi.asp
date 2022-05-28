<%
Dim objConn, stringSQL, strConnection, array_id, i, sql_id, id, nome, FSO, strPath
id = Request("radio")
'Caso ocorra algum erro os precessos não são interrompidos 
'e é passado para a próxima linha de comando
Set FSO= Server.CreateObject("Scripting.FileSystemObject")
Set objConn =  Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("../../../ch_database/bd_hermes.mdb")
stringSQL = "SELECT * FROM novidades WHERE id = "&id&""
Set ObjRs = objConn.Execute(stringSQL)


diretorio= ("../")
final = diretorio + objRs("novidades")
response.Write final

FSO.DeleteFile final, true




Set TB = Nothing
objConn.close
Set objConn = Nothing		  
Response.Redirect "action_excluircoluna.asp?id="&id
%>