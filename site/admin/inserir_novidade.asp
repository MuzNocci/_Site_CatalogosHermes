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
    <td width="7" background="../imagens/sombraesquerda.gif">&nbsp;</td>
    <td width="770" height="300" valign="middle" bgcolor="#FFFFFF">
<form action="action_inserirnov.asp" method="post" enctype="multipart/form-data" name="cadastro" id="cadastro"  onSubmit="return validaForm()">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
          		<tr>
      <td align="center" valign="top"> 
                <table width="770" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="50" colspan="2" align="left" width="770"><img src="imagens/inclusaocolunatit.gif" width="173" height="23"></td>
                  </tr>
                  <tr>
                    <td colspan="2" align="left" class="textoform" width="770">T&iacute;tulo da Novidade:</td>
</tr>
                  <tr>
                    <td colspan="2" align="left" width="770"><input name="titulo" type="text" class="form" size="60" maxlength="60"></td>
                  </tr>
                  <tr valign="bottom">
                    <td height="30" colspan="2" align="left" class="textoform" width="770">Texto da Novidade:</td>
                  </tr>
                  <tr valign="bottom">
                    <td colspan="2" align="left" class="textoform" width="770"><label>
                      <textarea name="novidade" cols="90" rows="6" class="formmsg" id="novidade" style="width: 770"></textarea>
                    </label></td>
                  </tr>
                  <tr valign="bottom">
                    <td height="30" colspan="2" align="left" class="textoform" width="770">Inserir a imagem: </td>
                  </tr>
                  <tr>
                    <td colspan="2" align="left" width="770">
                    <input name="FILE1" type="file" class="form" id="imagem" size="61"></td>
                  </tr>
                  <tr>
                    <td height="30" align="left" valign="top" width="681"><span class="style1">A imagem n&atilde;o deve ter mais do que 
                    300px de largura. </span></td>
                    <td align="left" width="89">&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="left" width="681"><div align="right">
                      <input name="Submit2" type="reset" class="form" value="Limpar">
&nbsp;                    </div></td>
                    <td align="left" width="89">&nbsp;
                    <input name="enviar" type="submit" class="form" id="enviar" value="Enviar">                    </td>
                  </tr>
                </table>                </td>
              </tr>
            </table>
      </form>
	</td>
    <td width="5" background="../imagens/sombradireita.gif">&nbsp;</td>
  </tr>
  </table>
</body>
</html>