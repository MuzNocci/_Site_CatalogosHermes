<% Server.ScriptTimeout = 300 %>
<%
Set Conexao = CreateObject("ADODB.CONNECTION")
conStr = "DBQ=" & Server.MapPath("../../../ch_database/bd_hermes.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}"
ConexaoAberta = FALSE
Sub AbreConexao()
	if not ConexaoAberta then
	Conexao.Open ConStr
	ConexaoAberta = True
	end if
end sub
Sub FechaConexao()
	if ConexaoAberta then
	Conexao.close
	ConexaoAberta = False
	end if
end sub 
%>