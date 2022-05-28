<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Server.ScriptTimeout = 300 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="restrito.asp"-->
<!--#include file="conectar.asp"-->
<%

Dim objConn, objRs, strQuery
Dim strConnection

'Conectando com o banco de dados contato.mdb
Set objConn =  Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("../../../ch_database/bd_hermes.mdb")
strQuery = "SELECT * FROM datas order by datas.ano, datas.mes DESC"
Set ObjRs = objConn.Execute(strQuery)
%>

<html>
<head>
<title>Administra&ccedil;&atilde;o</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../hermes.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin:0px; background-image: url('../imagens/fundo.jpg');
	
}
-->
</style>

</head>
<body>
<!--#include file="topbar.asp"-->
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="300" background="../imagens/sombraesquerda.gif">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
          <form method="get" action="action_excluirdata.asp">
          <tr>
            <td align="center" valign="top"> 
              <table width="770" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="50" align="left"><img src="imagens/excluirprofissional.gif" width="250" height="23"></td>
                </tr>
                <tr>
                  <td><table border="0" width="770" cellpadding="2" align="center">
                    <tr bgcolor="#000099">
                      <td width="454" bgcolor="#CC0000" class="titulocaixa"><div align="left">&nbsp;M&ecirc;s/Ano Cadastrado:</div></td>
                      <td width="54" height="2" align="center" bgcolor="#CC0000">
                        <input name="Submit" type="submit" class="form" value="Excluir"></td>
                      </tr>
                    <%While Not objRS.EOF %>
<tr bgcolor="#99CCFF">
                      <td bgcolor="#FFD9D9" class="admintextos"><%Response.write objRS("mes")%> 
                      / 
                        <%Response.write objRS("ANO")%></td>
                      <td width="54" height="2" align="center" bgcolor="#FFD9D9"><input type="radio" name="radio" value="<%=objRS(0)%>"></td>
                      </tr>
                    <%
  'Move para o pr&oacute;ximo registro
  objRS.MoveNext
  Wend
  'Fechando as conex&otilde;es
  objRs.close
  objConn.close
  Set objRs = Nothing
  Set objConn = Nothing
  %>
                    <tr bgcolor="#000099">
                      <td height="10" align="center" bgcolor="#CC0000"></td>
                      <td height="10" align="center" bgcolor="#CC0000"></td>
                      </tr>
                  </table></td>
                </tr>
            </table>                </td>
          </tr>
          </form>
        </table>
    </td>
  </tr>
</table>
</body>
</html>