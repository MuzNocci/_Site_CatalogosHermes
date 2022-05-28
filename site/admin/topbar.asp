
<head>
<style type="text/css">
<!--
body {
	margin:0px; background-image: url('../imagens/fundo.jpg');
	
}
-->
</style>
<script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
<script language="javascript">
navHover = function() {
	var lis = document.getElementById("navmenu").getElementsByTagName("LI");
	for (var i=0; i<lis.length; i++) {
		lis[i].onmouseover=function() {
			this.className+=" iehover";
		}
		lis[i].onmouseout=function() {
			this.className=this.className.replace(new RegExp(" iehover\\b"), "");
		}
	}
}
if (window.attachEvent) window.attachEvent("onload", navHover);
</script>
<link href="menu.css" rel="stylesheet" type="text/css" />
<link href="SpryAssets/SpryMenuBarHorizontal.css" rel="stylesheet" type="text/css" />
</head>

<table width="770" height="110" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="164" height="100"><img border="0" src="imagens/logo.png"></td>
    <td align="right"><font face="Arial" size="4" color="#CC0000">Administração<br>
	</font><font face="Arial" size="1" color="#CC0000"><br>
	<%
	Set objConn1 =  Server.CreateObject("ADODB.Connection")
	objConn1.Open "DBQ=" & Server.MapPath("../../../ch_database/bd_hermes.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","Admin",""
	Set dados = objConn1.execute("select * from contador where idcontador="&month(Now)&"")
	Application("acesso") = dados("acessos")
	Application("impressoes") = dados("impressoes")
	dados.close
	Set dados = nothing
	objConn1.close
	set objConn1 = nothing
	%></font><b><font face="Arial" style="font-size: 9pt">Informações de acessos do Mês<br>
	Acesso: <%=Application("acesso")%><br>
	Impressões de página: <%=Application("impressoes")%></font></b></td>
  </tr>
  <tr>
    <td width="758" height="20" colspan="2" bgcolor="#CC0000">
<ul id="MenuBar" class="MenuBarHorizontal">
      <li><b><font face="Arial" style="font-size: 9pt"><a href="default.asp">Home</a></font></b></li>
      <li><b><font face="Arial" style="font-size: 9pt"><a class="MenuBarItemSubmenu" href="#">Not&iacute;cias</a>
        </font></b>
        <ul>
          <li><b><font face="Arial" style="font-size: 9pt"><a href="inserir_noticia.asp">Inserir Not&iacute;cias</a></font></b></li>
        <li><b><font face="Arial" style="font-size: 9pt"><a href="selec_alterar.asp">Alterar Not&iacute;cias</a></font></b></li>
        <li><b><font face="Arial" style="font-size: 9pt"><a href="selec_apagar.asp">Apagar Not&iacute;cias</a></font></b></li>
        </ul>
      </li>
            <li><b><font face="Arial" style="font-size: 9pt"><a class="MenuBarItemSubmenu" href="#">Novidades</a>
        	</font></b>
        <ul>
          <li><b><font face="Arial" style="font-size: 9pt"><a href="inserir_novidade.asp">Inserir Novidades</a></font></b></li>
        <li><b><font face="Arial" style="font-size: 9pt"><a href="selec_apagarnov.asp">Apagar Novidades</a></font></b></li>
        </ul>
      </li>
       		<li><b><font face="Arial" style="font-size: 9pt"><a class="MenuBarItemSubmenu" href="#">Data de Pedidos</a>
        	</font></b>
        <ul>
          <li><b><font face="Arial" style="font-size: 9pt"><a href="inserir_datas.asp">Novas Datas</a></font></b></li>
        <li><b><font face="Arial" style="font-size: 9pt"><a href="selec_apagar_datas.asp">Apagar Datas</a></font></b></li>
        </ul>
      </li>
      </ul></td>
  </tr>
</table>

<script type="text/javascript">
<!--
var MenuBar1 = new Spry.Widget.MenuBar("MenuBar", {imgDown:"../../SpryAssets/SpryMenuBarDownHover.gif", imgRight:"../../SpryAssets/SpryMenuBarRightHover.gif"});
//-->
</script>