<% Server.ScriptTimeout = 300 %>
<%	if Session("login")="" or Session("login") = False then %>
<table border=0 width=100% height=65% cellspacing=5 cellpadding=5>
<tr>
<td width=80% align=center>
	<p><font color="#FF0000" face=verdana size=2><B>Voc� n�o tem permiss�o para acessar esta p�gina. <br>
	Identifique-se, usando seu login e sua senha.</B></FONT>
	<p><!--#include file="login.asp"--></td></tr>
	</table>
</table>
<%Response.End
End if%>










