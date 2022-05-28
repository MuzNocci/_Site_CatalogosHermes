<% Server.ScriptTimeout = 300 %>
<!--#include file="conectar.asp"-->
<%
'Declaro as variaveis
Dim sUsername, sPassword
'recuperar nossos valores do textbox do formulário e atribui as variáveis

sUsername=Request.Form("login")
sPassword=Request.Form("login")

'Chama a funcaun IllegalChars para checar os caracteres
If IllegalChars(sUsername)=True OR IllegalChars(sPassword)=True Then
Response.redirect("default.asp")
End If

Function IllegalChars(sInput)
'Declaro as variaveis
Dim sBadChars, iCounter
'Set IllegalChars para False
IllegalChars=False
'Crai um array de caracters ilegais e palavras
sBadChars=array("select", "drop", ";", "--", " insert", "delete", "xp_", _
"#", "%", "&", "'", "(", ")", "/", "\", ":", ";", "<", ">", "=", "[", "]", "?", "`", "|")
'Executamos os laços através dos sBadChars do array usando contador & função de UBound
For iCounter = 0 to uBound(sBadChars)
'Usa a Funcaun Instr para checar a presença de caracters ilegais na variavel
If Instr(sInput,sBadChars(iCounter))>0 Then
IllegalChars=True
End If
Next
End function
%>
<% AbreConexao
	Dim Sql	
	Dim RS	
	Sql = "SELECT * FROM Admin WHERE login = '" & Request.Form("login") & "' "
	Sql = Sql & "AND senha='" & Request.Form("senha") & "' "
	set Rs = Conexao.Execute(Sql)
	if not rs.eof then
	Session("login") = RS("login")
	Session("id_cod") = RS("id_cod")
	Response.Redirect "default.asp"
	else
	Response.Redirect "default.asp"
	end if
FechaConexao
%>