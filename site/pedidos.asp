<%
Function formata(frase)
	Temp = Split(LCase(frase)," ")
	For i = LBound(Temp) to UBound(Temp)
	    if Temp(i) = "dos" or Temp(i) = "da" or Temp(i) = "das" or Temp(i) = "de" or Temp(i) = "di" or Temp(i) = "do" then
            FraseTemp = FraseTemp &" "& Temp(i)
        else
            FraseTemp = FraseTemp &" "& UCase(Left(Temp(i),1)) & mid(Temp(i),2)
		end if
	Next
	formata = FraseTemp
End Function

Function fstring(str)
	U = 2 - len(str)
	fstring = Left(str, Abs(U))
End Function

nome = trim(formata(request.form("nome")))
nometitular = trim(formata(request.form("nometitular")))
endereco = trim(formata(request.form("endereco")))
referencia = trim(formata(request.form("referencia")))
sresidencia = request.form("sresidencia")
tresidencia = request.form("tresidencia")
nascimento = request.form("nascimento")
bairro = trim(formata(request.form("bairro")))
cidade = trim(formata(request.form("cidade")))
cep = request.form("cep")
estado = trim(request.form("estado"))
sexo = request.form("sexo")
tel = request.form("tel")
cel = request.form("cel")
telc = request.form("telc")
email = lcase(request.form("email"))
civil = request.form("civil")
cpf =  request.form("cpf")
id =  request.form("id")
nmae =  trim(formata(request.form("nmae")))
dmae =  request.form("dmae")
npai =  trim(formata(request.form("npai")))
dpai =  request.form("dpai")
orgao =  request.form("orgao")
revendooutros =  request.form("revendooutros")
nfilho1 =  trim(formata(request.form("nfilho1")))
nfilho2 =  trim(formata(request.form("nfilho2")))
nfilho3 =  trim(formata(request.form("nfilho3")))
nfilho4 =  trim(formata(request.form("nfilho4")))
dfilho1 =  request.form("dfilho1")
dfilho2 =  request.form("dfilho2")
dfilho3 =  request.form("dfilho3")
dfilho4 =  request.form("dfilho4")
formapagamento = request.form("FP")
codigo = request.form("codigo")
numerocard = request.form("numerocard")
codigoseg = request.form("codigoseg")
mes = request.form("mes")
assunto = request.form("assunto")
mensagem = request.form("mensagem")
ano = request.form("ano")
parcelas = request.form("parcelas")
if parcelas <> "" then
parcelas = cdbl(parcelas)
end if
cartao = request.form("pagamento")
codigopvpv = request.form("codigopvpv")
produtopvpv = request.form("produtopvpv")
str_qtd = "u[]"
str_referencia = "u[]"
str_tamanho = "u[]"
str_descricao = "u[]"
str_pagina = "u[]"
str_cor = "u[]"
str_unitario = "u[]"
str_total = "u[]"
for i = 1 to 100
str_qtd = str_qtd & request.form("str_qtd"& i) & "[]"
str_referencia = str_referencia & request.form("str_referencia"& i) & "[]"
str_tamanho = str_tamanho & request.form("str_tamanho"& i) & "[]"
str_descricao = str_descricao & request.form("str_descricao"& i) & "[]"
str_pagina = str_pagina & request.form("str_pagina"& i) & "[]"
str_cor = str_cor & request.form("str_cor"& i) & "[]"
str_unitario = str_unitario & request.form("str_unitario"& i) & "[]"
str_total = str_total & request.form("str_total"& i) & "[]"
next

vquant = split(fstring(str_qtd),"[]")
vref = split(fstring(str_referencia),"[]")
vtam = split(fstring(str_tamanho),"[]")
vdesc = split(fstring(str_descricao),"[]")
vpag = split(fstring(str_pagina),"[]")
vcor = split(fstring(str_cor),"[]")
vunit = split(fstring(str_unitario),"[]")
vtotal = split(fstring(str_total),"[]")
if not IsArray(vquant) then
    vquant = array(vquant)
	vref = array(vref)
	vtam = array(vtam)
	vdesc = array(vdesc)
	vpag = array(vpag)
	vcor = array(vcor)
    vunit = array(vunit)
	vtotal = array(vtotal)
end if
	htmlemail = htmlemail & "<br /><font face=Verdana color=#000000><strong>Informações do Solicitante:</strong></font>"
	htmlemail = htmlemail & "<BR><BR><font color=#000000 face=Verdana, Arial size=2><strong>Nome:</strong></font><br>"
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & nome & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Endereço: </strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & endereco & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Bairro:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & bairro & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Cidade:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & cidade & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>E-mail:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & email & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Telefone:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & tel & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Celular:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & cel & "</font><br><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Forma de Pagamento:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & cartao & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Número do Cartão:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & numerocard & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Cod. de Seguranca:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & codigoseg & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Data de Validade:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & mes & "/" & ano & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Número de Parcelas:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & parcelas & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Nome da Cliente:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & nometitular & "</font><br>"		
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Telefone da Cliente:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & telc & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Codigo de Troca PVPV:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & codigopvpv & "</font><br>"
	htmlemail = htmlemail & "<font color=#000000 face=Verdana, Arial size=2><strong>Produto de Troca PVPV:</strong></font><br>"	
	htmlemail = htmlemail & "<font color=#333333 face=Verdana, Arial size=2>" & produtopvpv & "</font><br><br>"
	htmlemail = htmlemail & "<table width=800 border=0 cellspacing=0 cellpadding=0>"
	htmlemail = htmlemail & "<tr><td colspan=6 align=center>FORMUL&Aacute;RIO DE PEDIDOS DO SITE</td></tr>"
	htmlemail = htmlemail & "<tr class=textonoticias2><td align=center bgcolor=#CCCCCC>Quant.</td><td align=center bgcolor=#CCCCCC>Refer&ecirc;ncia</td>"
    htmlemail = htmlemail & "<td align=center bgcolor=#CCCCCC>Tamanho</td><td align=center bgcolor=#CCCCCC>Descrição</td><td align=center bgcolor=#CCCCCC>P&aacute;gina</td><td align=center bgcolor=#CCCCCC>Cor</td>"
    htmlemail = htmlemail & "<td align=center bgcolor=#CCCCCC>Pre&ccedil;o Unit&aacute;rio</td><td align=center bgcolor=#CCCCCC>Pre&ccedil;o Total</td></tr><br /><br />"		
for u = 1 to ubound(vquant)
	htmlemail = htmlemail & "<tr><td align=center>" & vquant(u) &"</td><td align=center>" & vref(u) & "</td><td align=center>" & vdesc(u) & "</td><td align=center>" & vtam(u) & "</td><td align=center>" & vpag(u) & "</td><td align=center>" & vcor(u) & "</td><td align=center>" & vunit(u) & "</td><td align=center>" & vtotal(u) & "</td></tr>"
next
	htmlemail = htmlemail & "</table><br /><br />"
Set oMail = Server.CreateObject("Persits.MailSender")
oMail.Host = "smtp.catalogoshermes.com.br"
oMail.Port = 587
oMail.UserName = "admin@catalogoshermes.com.br"
oMail.PassWord = "art@hermes123"
oMail.From = "hermes@catalogoshermes.com.br"
oMail.FromName = nome
oMail.AddAddress "hermes@catalogoshermes.com.br", "Catálogos Hermes"
oMail.AddReplyTo email, nome
oMail.Subject = "Passar Pedido"
oMail.Body = htmlemail
oMail.IsHTML = True
oMail.Send
If Err <> 0 Then
	erro = "<b><font color='red'> Erro ao enviar a mensagem.</font></b><br>"
	erro = erro & "<b>Erro.Description:</b> " & Err.Description & "<br>"
	erro = erro & "<b>Erro.Number:</b> "      & Err.Number & "<br>"
	erro = erro & "<b>Erro.Source:</b> "      & Err.Source & "<br>"
	response.write erro
Else
    response.redirect("http://www.catalogoshermes.com.br/site/?hl=pedidos.ok")
End If
%>