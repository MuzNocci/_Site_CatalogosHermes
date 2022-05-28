<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Server.ScriptTimeout = 300 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="restrito.asp"-->
<!--#include file="conectar.asp"-->
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
.style1 {
	font-size: 9px;
	font-family: "Trebuchet MS", Verdana, Arial;
	color: #990000;
}
-->
</style>
<script language="JavaScript">
	function validaForm(){
		//validar nome
		d = document.cadastro;
		if (d.titulo.value == ""){
			alert("Preencha o Título da Notícia!");
			d.titulo.focus();
			return false;
		}
		//validar telefone
		if (d.noticia1.value == ""){
			alert("Preencha pelo menos o 1º Parágrafo!");
			d.noticia1.focus();
			return false;
		}
				
}
</script>
      <script language="VBScript">
SUB enviar_ONCLICK()
  cadastro.enviar.Value = "Aguarde enviando..."
END SUB
</script>
</head>
<body>
<!--#include file="topbar.asp"-->
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="770" height="8" colspan="3" align="center" valign="top">
      <table width="100" border="0" cellspacing="0" cellpadding="0" height="100%">
      <form action="action_inserirnoticia.asp" method="post" name="cadastro" id="cadastro"  onSubmit="return validaForm()">
        <tr>
          <td align="center" valign="top">
          <table width="770" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="50" colspan="2" align="left" width="770"><img src="imagens/inclusaodenticiastit.gif" width="222" height="22"></td>
            </tr>
            <tr>
              <td colspan="2" align="left" class="textoform" width="770">T&iacute;tulo da Not&iacute;cias:</td>
            </tr>
            <tr>
              <td colspan="2" align="left" width="770"><input name="titulo" type="text" class="form" size="60" maxlength="60"></td>
            </tr>
            <tr>
              <td height="30" colspan="2" align="left" valign="bottom" class="textoform" width="770">Corpo da Not&iacute;cia: </td>
            </tr>
            <tr>
              <td colspan="2" align="left" width="770">
              <textarea name="noticia" cols="90" rows="6" class="formmsg" id="noticia" style="width: 770">
    </textarea></td>
            </tr>
            <tr>
              <td height="60" align="left" width="697"><div align="right">
                <input name="Submit2" type="reset" class="form" value="Limpar">
                &nbsp; </div></td>
              <td align="left" width="73">&nbsp;
                <input name="enviar" type="submit" class="form" id="enviar" value="Enviar"></td>
            </tr>
          </table></td>
        </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
</body>
</html>