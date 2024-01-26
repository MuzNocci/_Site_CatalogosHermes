<%
StrTopo = "" & vbNewline
StrTopo = StrTopo & "<!--" & vbNewline
StrTopo = StrTopo & "###########################################################" & vbNewline
StrTopo = StrTopo & "                                                           " & vbNewline
StrTopo = StrTopo & "                                     CATALOGOSHERMES.COM.BR" & vbNewline
StrTopo = StrTopo & "                          CONCEIVED BY: DISTRIBUIDOR HERMES" & vbNewline
StrTopo = StrTopo & "                                                           " & vbNewline
StrTopo = StrTopo & "                              DEVELOPED BY: ARTLIZANDO LTDA" & vbNewline
StrTopo = StrTopo & "                     Campos dos Goytacazes - Rio de Janeiro" & vbNewline
StrTopo = StrTopo & "                   Fone: (22) 27252110 ou TIM (22) 81146861" & vbNewline
StrTopo = StrTopo & "             Copyright(c) "&year(Now)&" Todos os direitos reservados" & vbNewline
StrTopo = StrTopo & "                                                           " & vbNewline
StrTopo = StrTopo & "###########################################################" & vbNewline
StrTopo = StrTopo & "-->" & vbNewline
StrTopo = StrTopo & "" & vbNewline

response.write StrTopo

'String Segurança
Application("fwsseg") = "fwshermeshagf87agsfqgwrgv67g"

'Funções
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

'Applications
Application("hlpage") = request.querystring("hl")
Application("action") = request.querystring("action")
if Application("hlpage") = "" and Application("action") = "" then
Application("hlpage") = "inicio"
end if
Application("erroconfirm") = ""

'Variáveis
nome = trim(formata(request.querystring("nome")))
nometitular = trim(formata(request.querystring("nometitular")))
endereco = trim(formata(request.querystring("endereco")))
referencia = trim(formata(request.querystring("referencia")))
sresidencia = request.querystring("sresidencia")
tresidencia = request.querystring("tresidencia")
nascimento = request.querystring("nascimento")
bairro = trim(formata(request.querystring("bairro")))
cidade = trim(formata(request.querystring("cidade")))
cep = request.querystring("cep")
estado = trim(request.querystring("estado"))
sexo = request.querystring("sexo")
tel = request.querystring("tel")
cel = request.querystring("cel")
telc = request.querystring("telc")
email = lcase(request.querystring("email"))
civil = request.querystring("civil")
cpf =  request.querystring("cpf")
id =  request.querystring("id")
nmae =  trim(formata(request.querystring("nmae")))
dmae =  request.querystring("dmae")
npai =  trim(formata(request.querystring("npai")))
dpai =  request.querystring("dpai")
orgao =  request.querystring("orgao")
revendooutros =  request.querystring("revendooutros")
nfilho1 =  trim(formata(request.querystring("nfilho1")))
nfilho2 =  trim(formata(request.querystring("nfilho2")))
nfilho3 =  trim(formata(request.querystring("nfilho3")))
nfilho4 =  trim(formata(request.querystring("nfilho4")))
dfilho1 =  request.querystring("dfilho1")
dfilho2 =  request.querystring("dfilho2")
dfilho3 =  request.querystring("dfilho3")
dfilho4 =  request.querystring("dfilho4")
formapagamento = request.querystring("FP")
codigo = request.querystring("codigo")
numerocard = request.querystring("numerocard")
codigoseg = request.querystring("codigoseg")
mes = request.querystring("mes")
assunto = request.querystring("assunto")
mensagem = request.querystring("mensagem")
ano = request.querystring("ano")
parcelas = request.querystring("parcelas")
if parcelas <> "" then
parcelas = cdbl(parcelas)
end if
cartao = request.querystring("pagamento")
codigopvpv = request.querystring("codigopvpv")
produtopvpv = request.querystring("produtopvpv")
for i = 1 to 100
b = request.querystring("str_total"& i)
if b <> "" and b <> "0,00" then
str_qtd = str_qtd & request.querystring("str_qtd"& i) & "[]"
str_referencia = str_referencia & request.querystring("str_referencia"& i) & "[]"
str_tamanho = str_tamanho & request.querystring("str_tamanho"& i) & "[]"
str_descricao = str_descricao & request.querystring("str_descricao"& i) & "[]"
str_pagina = str_pagina & request.querystring("str_pagina"& i) & "[]"
str_cor = str_cor & request.querystring("str_cor"& i) & "[]"
str_unitario = str_unitario & request.querystring("str_unitario"& i) & "[]"
str_total = str_total & request.querystring("str_total"& i) & "[]"
end if
next

'Ações

Select case Application("action")

case "cadastrar"
if nome = "" or len(nome) < 4 then
if nome = "" then
session("erro_nome") = "Digite seu nome!"
Application("erroconfirm") = "erro"
else
session("erro_nome") = "Nome inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_nome") = ""
end if
if endereco = "" or len(endereco) < 5 then
if endereco = "" then
session("erro_endereco") = "Digite o endereço!"
Application("erroconfirm") = "erro"
else
session("erro_endereco") = "Endereço inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_endereco") = ""
end if
if bairro = "" or len(bairro) < 2 then
if bairro = "" then
session("erro_bairro") = "Digite o bairro!"
Application("erroconfirm") = "erro"
else
session("erro_bairro") = "Bairro inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_bairro") = ""
end if
if cep = "" or len(cep) <> 8 then
if cep = "" then
session("erro_cep") = "Digite o CEP!"
Application("erroconfirm") = "erro"
else
session("erro_cep") = "CEP inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_cep") = ""
end if
if cidade = "" or len(cidade) < 2 then
if cidade = "" then
session("erro_cidade") = "Digite a cidade!"
Application("erroconfirm") = "erro"
else
session("erro_cidade") = "Cidade inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_cidade") = ""
end if
if tel = "" or len(tel) < 10 or isnumeric(tel) = false then
if tel = "" then
session("erro_tel") = "Digite o número do seu telefone!"
Application("erroconfirm") = "erro"
else
session("erro_tel") = "Telefone inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_tel") = ""
end if
if email = "" or instr(email,"@") < 1 or instr(email,".com") < 1 and instr(email,".net") < 1 and instr(email,".org") < 1 and instr(email,".info") < 1 and instr(email,".gov") < 1 then
if email = "" then
session("erro_email") = "Digite o e-mail!"
Application("erroconfirm") = "erro"
else
session("erro_email") = "E-mail inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_email") = ""
end if
strCpf = cpf
ElCpf  = cpf
s="" 
for x=1 to len(strCpf)
ch=mid(strCpf,x,1)
if asc(ch)>=48 and asc(ch)<=57 then
s=s & ch
end if
next
strCpf = s
Dim Numero(11), soma, resultado1, resultado2
if strCpf = "" then
strCpf = "inv"
end if
if len(strCpf) <> 11 then
strCpf = "inv"
elseif strCpf = "00000000000" then
strCpf = "inv"
elseif strCpf = "11111111111" then
strCpf = "inv"
elseif strCpf = "01234567890" then
strCpf = "inv"
elseif strCpf = "22222222222" then
strCpf = "inv"
elseif strCpf = "33333333333" then
strCpf = "inv"
elseif strCpf = "44444444444" then
strCpf = "inv"
elseif strCpf = "55555555555" then
strCpf = "inv"
elseif strCpf = "66666666666" then
strCpf = "inv"
elseif strCpf = "77777777777" then
strCpf = "inv"
elseif strCpf = "88888888888" then
strCpf = "inv"
elseif strCpf = "99999999999" then
strCpf = "inv"
elseif strCpf = "12345678909" then
strCpf = "inv"
else
Numero(1) = Cint(Mid(strCpf,1,1))
Numero(2) = Cint(Mid(strCpf,2,1))
Numero(3) = Cint(Mid(strCpf,3,1))
Numero(4) = Cint(Mid(strCpf,4,1))
Numero(5) = Cint(Mid(strCpf,5,1))
Numero(6) = CInt(Mid(strCpf,6,1))
Numero(7) = Cint(Mid(strCpf,7,1))
Numero(8) = Cint(Mid(strCpf,8,1))
Numero(9) = Cint(Mid(strCpf,9,1))
Numero(10) = Cint(Mid(strCpf,10,1))
Numero(11) = Cint(Mid(strCpf,11,1))
soma = 10 * Numero(1) + 9 * Numero(2) + 8 * Numero(3) + 7 * Numero(4) + 6 * Numero(5) + 5 * Numero(6) + 4 * Numero(7) + 3 * Numero(8) + 2 * Numero(9)
soma = soma -(11 * (int(soma / 11)))
if soma = 0 or soma = 1 then
resultado1 = 0
else
resultado1 = 11 - soma
end if
if resultado1 = Numero(10) then
soma = Numero(1) * 11 + Numero(2) * 10 + Numero(3) * 9 + Numero(4) * 8 + Numero(5) * 7 + Numero(6) * 6 + Numero(7) * 5 + Numero(8) * 4 + Numero(9) * 3 + Numero(10) * 2
soma = soma -(11 * (int(soma / 11)))
if soma = 0 or soma = 1 then
resultado2 = 0
else
resultado2 = 11 - soma
end if
if resultado2 = Numero(11) then
else
strCpf = "inv"
end if
else 
strCpf = "inv"
end if
end if
if strCpf = "inv" then
strCpf = ElCpf
end if
s="" 
for x=1 to len(ElCpf)
ch=mid(ElCpf,x,1)
if asc(ch)>=48 and asc(ch)<=57 then
s=s & ch
end if
next
ElCpf = s
if len(ElCpf) <> 11 then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "00000000000" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "11111111111" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "01234567890" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "22222222222" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "33333333333" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "44444444444" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "55555555555" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "66666666666" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "77777777777" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "88888888888" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "99999999999" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
elseif ElCpf = "12345678909" then
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
else
Numero(1) = Cint(Mid(ElCpf,1,1))
Numero(2) = Cint(Mid(ElCpf,2,1))
Numero(3) = Cint(Mid(ElCpf,3,1))
Numero(4) = Cint(Mid(ElCpf,4,1))
Numero(5) = Cint(Mid(ElCpf,5,1))
Numero(6) = CInt(Mid(ElCpf,6,1))
Numero(7) = Cint(Mid(ElCpf,7,1))
Numero(8) = Cint(Mid(ElCpf,8,1))
Numero(9) = Cint(Mid(ElCpf,9,1))
Numero(10) = Cint(Mid(ElCpf,10,1))
Numero(11) = Cint(Mid(ElCpf,11,1))
soma = 10 * Numero(1) + 9 * Numero(2) + 8 * Numero(3) + 7 * Numero(4) + 6 * Numero(5) + 5 * Numero(6) + 4 * Numero(7) + 3 * Numero(8) + 2 * Numero(9)
soma = soma -(11 * (int(soma / 11)))
if soma = 0 or soma = 1 then
resultado1 = 0
else
resultado1 = 11 - soma
end if
if resultado1 = Numero(10) then
soma = Numero(1) * 11 + Numero(2) * 10 + Numero(3) * 9 + Numero(4) * 8 + Numero(5) * 7 + Numero(6) * 6 + Numero(7) * 5 + Numero(8) * 4 + Numero(9) * 3 + Numero(10) * 2
soma = soma -(11 * (int(soma / 11)))
if soma = 0 or soma = 1 then
resultado2 = 0
else
resultado2 = 11 - soma
end if
if resultado2 = Numero(11) then
session("erro_cpf") = ""
else
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
end if
else 
session("erro_cpf") = "CPF inválido!"
Application("erroconfirm") = "erro"
end if
end if
if ElCpf = "" then
session("erro_cpf") = "Digite o número do seu CPF!"
Application("erroconfirm") = "erro"
end if
if id = "" or len(id) < 7 then
if id = "" then
session("erro_id") = "Digite o número do seu RG!"
Application("erroconfirm") = "erro"
else
session("erro_id") = "RG inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_id") = ""
end if
if Application("erroconfirm") = "erro" then
response.redirect("?hl=cadastre-se&nome="&nome&"&endereco="&endereco&"&referencia="&referencia&"&bairro="&bairro&"&cep="&cep&"&cidade="&cidade&"&estado="&estado&"&sexo="&sexo&"&tel="&tel&"&cel="&cel&"&email="&email&"&civil="&civil&"&nascimento="&nascimento&"&cpf="&cpf&"&id="&id&"&orgao="&orgao&"&tresidencia="&tresidencia&"&sresidencia="&sresidencia&"&nmae="&nmae&"&dmae="&dmae&"&npai="&npai&"&dpai="&dpai&"&nfilho1="&nfilho1&"&dfilho1="&dfilho1&"&nfilho2="&nfilho2&"&dfilho2="&dfilho2&"&nfilho3="&nfilho3&"&dfilho3="&dfilho3&"&nfilho4="&nfilho4&"&dfilho4="&dfilho4&"&revendooutros="&revendooutros&"")
end if
htmlemail = htmlemail & "<font face=Verdana color=""#000000""><strong>Cadastramento do Site:</strong></font>"
htmlemail = htmlemail & "<BR><BR><font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Nome:</strong></font><br>"
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & nome & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Endereço:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & endereco & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Referência para Entrega:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & referencia & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Bairro:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & bairro & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Cidade:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & cidade & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Estado:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & trim(formata(estado)) & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Data de Nascimento:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & nascimento & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>CEP:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & cep & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Sexo:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & sexo & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Telefone:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>(" & left(tel,2) & ") "&right(tel,8)&"</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Celular:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>(" & left(cel,2) & ") "&right(cel,8)&"</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>E-mail:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & email & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Estado Civíl:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & civil & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>CPF:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & cpf & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Identidade:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & id & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Órgão Expedidor:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & orgao & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Tempo de Residência:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & tresidencia & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Situação da Residência:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & sresidencia & "</strong></font><br>"     
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Nome da Mãe:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & nmae & "</strong></font><br>"     
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Data de Nascimento da Mãe:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & dmae & "</strong></font><br>"     
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Nome do Pai:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & npai & "</strong></font><br>"     
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Data de Nascimento do Pai:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & dpai & "</strong></font><br>"     
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Nome e Data de Nascimento dos Filhos menores de 14 anos:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & nfilho1 & ""& dfilho1 & "</strong></font><br>"   
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & nfilho2 & ""& dfilho2 & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & nfilho3 & ""& dfilho3 & "</strong></font><br>"   		   
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & nfilho4 & ""& dfilho4 & "</strong></font><br>"   	
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Revendo Outras Revistas:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & revendooutros & "</strong></font><br>"     			
Set oMail = Server.CreateObject("Persits.MailSender")
oMail.Host = "smtp.catalogoshermes.com.br"
oMail.Port = 587
oMail.UserName = "hermes@catalogoshermes.com.br"
oMail.PassWord = "hermes123"
oMail.From = "hermes@catalogoshermes.com.br"
oMail.FromName = nome
oMail.AddAddress "hermes@catalogoshermes.com.br", "Cadastro de revendedores"
oMail.AddReplyTo email, nome
oMail.Subject = "Cadastro de revendedor(a)"
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
    response.redirect("?hl=cadastrar.ok")
End If

case "enviarpedido_2"
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
for i = 0 to ubound(vquant)
	htmlemail = htmlemail & "<tr><td align=center>" & vquant(i) &"</td><td align=center>" & vref(i) & "</td><td align=center>" & vdesc(i) & "</td><td align=center>" & vtam(i) & "</td><td align=center>" & vpag(i) & "</td><td align=center>" & vcor(i) & "</td><td align=center>" & vunit(i) & "</td><td align=center>" & vtotal(i) & "</td></tr>"
next
	htmlemail = htmlemail & "</table><br /><br />"
Set oMail = Server.CreateObject("Persits.MailSender")
oMail.Host = "smtp.catalogoshermes.com.br"
oMail.Port = 587
oMail.UserName = "hermes@catalogoshermes.com.br"
oMail.PassWord = "hermes123"
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
    response.redirect("?hl=pedidos.ok")
End If

case "catalogos"
if nome = "" or len(nome) < 4 then
if nome = "" then
session("erro_nome") = "Digite seu nome!"
Application("erroconfirm") = "erro"
else
session("erro_nome") = "Nome inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_nome") = ""
end if
if endereco = "" or len(endereco) < 5 then
if endereco = "" then
session("erro_endereco") = "Digite o endereço!"
Application("erroconfirm") = "erro"
else
session("erro_endereco") = "Endereço inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_endereco") = ""
end if
if bairro = "" or len(bairro) < 2 then
if bairro = "" then
session("erro_bairro") = "Digite o bairro!"
Application("erroconfirm") = "erro"
else
session("erro_bairro") = "Bairro inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_bairro") = ""
end if
if cep = "" or len(cep) <> 8 or isnumeric(cep) = false then
if cep = "" then
session("erro_cep") = "Digite o CEP!"
Application("erroconfirm") = "erro"
else
session("erro_cep") = "CEP inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_cep") = ""
end if
if cidade = "" or len(cidade) < 2 then
if cidade = "" then
session("erro_cidade") = "Digite a cidade!"
Application("erroconfirm") = "erro"
else
session("erro_cidade") = "Cidade inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_cidade") = ""
end if
if tel = "" or len(tel) < 10 or isnumeric(tel) = false then
if tel = "" then
session("erro_tel") = "Digite o número do seu telefone!"
Application("erroconfirm") = "erro"
else
session("erro_tel") = "Telefone inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_tel") = ""
end if
if email = "" or instr(email,"@") < 1 or instr(email,".com") < 1 and instr(email,".net") < 1 and instr(email,".org") < 1 and instr(email,".info") < 1 and instr(email,".gov") < 1 then
if email = "" then
session("erro_email") = "Digite o e-mail!"
Application("erroconfirm") = "erro"
else
session("erro_email") = "E-mail inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_email") = ""
end if
if Application("erroconfirm") = "erro" then
response.redirect("?hl=pecacatalogos&nome="&nome&"&endereco="&endereco&"&bairro="&bairro&"&cep="&cep&"&cidade="&cidade&"&estado="&estado&"&tel="&tel&"&cel="&cel&"&email="&email&"")
end if
htmlemail = htmlemail & "<font face=Verdana color=""#000000""><strong>Cadastramento do Site:</strong></font>"
htmlemail = htmlemail & "<BR><BR><font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Nome:</strong></font><br>"
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & nome & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Endereço:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & endereco & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Bairro:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & bairro & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Cidade:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & cidade & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Estado:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & trim(formata(estado)) & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>CEP:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & cep & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Telefone:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>(" & left(tel,2) & ") "&right(tel,8)&"</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Celular:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>(" & left(cel,2) & ") "&right(cel,8)&"</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>E-mail:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & email & "</strong></font><br>"
Set oMail = Server.CreateObject("Persits.MailSender")
oMail.Host = "smtp.catalogoshermes.com.br"
oMail.Port = 587
oMail.UserName = "hermes@catalogoshermes.com.br"
oMail.PassWord = "hermes123"
oMail.From = "hermes@catalogoshermes.com.br"
oMail.FromName = nome
oMail.AddAddress "hermes@catalogoshermes.com.br", "Catálogos Hermes"
oMail.AddReplyTo email, nome
oMail.Subject = "Pedido de catálogo"
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
    response.redirect("?hl=catalogos.ok")
End If

case "contact"
if nome = "" or len(nome) < 4 then
if nome = "" then
session("erro_nomec") = "Digite seu nome!"
Application("erroconfirm") = "erro"
else
session("erro_nomec") = "Nome inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_nomec") = ""
end if
if cidade = "" or len(cidade) < 2 then
if cidade = "" then
session("erro_cidadec") = "Digite a cidade!"
Application("erroconfirm") = "erro"
else
session("erro_cidadec") = "Cidade inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_cidadec") = ""
end if
if tel = "" or len(tel) < 10 or isnumeric(tel) = false then
if tel = "" then
session("erro_telc") = "Digite o número do seu telefone!"
Application("erroconfirm") = "erro"
else
session("erro_telc") = "Telefone inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_telc") = ""
end if
if email = "" or instr(email,"@") < 1 or instr(email,".com") < 1 and instr(email,".net") < 1 and instr(email,".org") < 1 and instr(email,".info") < 1 and instr(email,".gov") < 1 then
if email = "" then
session("erro_emailc") = "Digite o e-mail!"
Application("erroconfirm") = "erro"
else
session("erro_emailc") = "E-mail inválido!"
Application("erroconfirm") = "erro"
end if
else
session("erro_emailc") = ""
end if
if mensagem = "" or len(mensagem) < 2 then
if mensagem = "" then
session("erro_mensagemc") = "Digite a mensagem!"
Application("erroconfirm") = "erro"
else
session("erro_mensagemc") = "A mensagem deve conter no mínimo 2 caractéres!"
Application("erroconfirm") = "erro"
end if
else
session("erro_mensagemc") = ""
end if
if Application("erroconfirm") = "erro" then
response.redirect("?hl=contato&nome="&nome&"&cidade="&cidade&"&estado="&estado&"&tel="&tel&"&cel="&cel&"&email="&email&"&assunto="&assunto&"&mensagem="&mensagem&"")
end if
htmlemail = htmlemail & "<font face=Verdana color=""#000000""><strong>Cadastramento do Site:</strong></font>"
htmlemail = htmlemail & "<BR><BR><font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Nome:</strong></font><br>"
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & nome & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Cidade:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & cidade & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Estado:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & trim(formata(estado)) & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Telefone:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>(" & left(tel,2) & ") "&right(tel,8)&"</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Celular:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>(" & left(cel,2) & ") "&right(cel,8)&"</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>E-mail:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & email & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Assunto:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & assunto & "</strong></font><br>"
htmlemail = htmlemail & "<font color=""#000000"" face=""Verdana, Arial"" size=""2""><strong>Mensagem:</strong></font><br>"	
htmlemail = htmlemail & "<font color=""#333333"" face=""Verdana, Arial"" size=""2""><strong>" & mensagem & "</strong></font><br>"
Set oMail = Server.CreateObject("Persits.MailSender")
oMail.Host = "smtp.catalogoshermes.com.br"
oMail.Port = 587
oMail.UserName = "hermes@catalogoshermes.com.br"
oMail.PassWord = "hermes123"
oMail.From = "hermes@catalogoshermes.com.br"
oMail.FromName = nome
oMail.AddAddress "hermes@catalogoshermes.com.br", "Contato via site"
oMail.AddReplyTo email, nome
oMail.Subject = assunto
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
    response.redirect("?hl=contato.ok")
End If

end select
%><!-- #include file="includes/conect.catalogoshermes.asp" --><%

Set cont = objconn.execute("select * from contador where idcontador="&month(now)&"")
if not cont.eof then
Application("chacessos") = cont("acessos")*1
Application("chimpressos") = cont("impressoes")*1
Application("mes") = month(now)
if Application("chacessos") = 0 and Application("chimpressos") = 0 or cont("ano") <> year(now) then
Set contup = objconn.execute("update contador set acessos=1, impressoes=1, ano="&year(now)&" where idcontador="&Application("mes")&"")
Session("xacessos") = "ok"
else
if Session("xacessos") <> "ok" then
Set contup = objconn.execute("update contador set acessos="&Application("chacessos")+1&",impressoes="&Application("chimpressos")+1&" where idcontador="&Application("mes")&"")
Session("xacessos") = "ok"
else
Set contup = objconn.execute("update contador set impressoes="&Application("chimpressos")+1&" where idcontador="&Application("mes")&"")
end if
end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="pt-br" lang="pt-br">
<head>
<meta http-equiv="Content-Language" content="pt-br">
<title>Catálogos Hermes | Um shopping em suas mãos!</title>
<meta name="description" content="" />
<meta name="keywords" content="" />
<meta name="revisit-after" content="1" />
<meta name="autor" content="ARTlizando LTDA" />
<meta name="company" content="ARTlizando LTDA" />
<link rev="made" href="mailto:atendimento@artlizando.com.br" />
<link rel="stylesheet" type="text/css" href="estilos/estilos.css" />
<!--[if IE]>
<link rel="stylesheet" type="text/css" href="estilos/estilos_ie.css" />
<![endif]-->
<link rel="shortcut icon" href="imagens/favicon.png" />
<link rel="icon" href="imagens/favicon.png" />
<meta content="NO-CACHE" http-equiv="pragma"></meta>
<meta content="no-cache" http-equiv="Cache-control"></meta>
<meta content="must-revalidate" http-equiv="Cache-control"></meta>
<meta content="max-age=0" http-equiv="Cache-control"></meta>
<meta content="Mon, 04 Dec 1999 21:29:02 GMT" http-equiv="Expires"></meta> 
<meta content="document" name="resource-type"></meta>
<meta content="ALL" name="robots"></meta>
<meta content="Global" name="distribution"></meta>
<meta content="General" name="rating"></meta>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
<script src="scripts/jquery-1.4.2.js" type="text/javascript"></script>
<script src="scripts/jquery.jcarousel.js" type="text/javascript"></script>
<script src="scripts/functions.js" type="text/javascript"></script>
<script language="JavaScript" type="text/javascript">
<!--
function Mascara(objeto, evt, mask) {
 
var LetrasU = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
var LetrasL = 'abcdefghijklmnopqrstuvwxyz';
var Letras  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
var Numeros = '0123456789';
var Fixos  = '().-:/ '; 
var Charset = " !\"#$%&\'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_/`abcdefghijklmnopqrstuvwxyz{|}~";

evt = (evt) ? evt : (window.event) ? window.event : "";
var value = objeto.value;
if (evt) {
 var ntecla = (evt.which) ? evt.which : evt.keyCode;
 tecla = Charset.substr(ntecla - 32, 1);
 if (ntecla < 32) return true;

 var tamanho = value.length;
 if (tamanho >= mask.length) return false;

 var pos = mask.substr(tamanho,1); 
 while (Fixos.indexOf(pos) != -1) {
  value += pos;
  tamanho = value.length;
  if (tamanho >= mask.length) return false;
  pos = mask.substr(tamanho,1);
 }

 switch (pos) {
   case '#' : if (Numeros.indexOf(tecla) == -1) return false; break;
   case 'A' : if (LetrasU.indexOf(tecla) == -1) return false; break;
   case 'a' : if (LetrasL.indexOf(tecla) == -1) return false; break;
   case 'Z' : if (Letras.indexOf(tecla) == -1) return false; break;
   case '*' : objeto.value = value; return true; break;
   default : return false; break;
 }
}
objeto.value = value; 
return true;
}
function MaskTelefone(objeto, evt) { 
return Mascara(objeto, evt, '(##) ####-####');
}
function MaskCPF(objeto, evt) { 
return Mascara(objeto, evt, '###.###.###-##');
}
function Referencia(objeto, evt) { 
return Mascara(objeto, evt, '##.###.##');
}
function Cor(objeto, evt) { 
return Mascara(objeto, evt, 'ZZZZZZZZ');
}
function Nquant(objeto, evt) { 
return Mascara(objeto, evt, '##');
}
function Ncartao(objeto, evt) { 
return Mascara(objeto, evt, '#### #### #### ####');
}
function Nseg(objeto, evt) { 
return Mascara(objeto, evt, '####');
}
function Nano(objeto, evt) { 
return Mascara(objeto, evt, '####');
}

function MascaraMoeda(objTextBox, SeparadorMilesimo, SeparadorDecimal, e){
    var SeparadorDecimal = ".";
    var SeparadorMilesimo = ".";
 var sep = 0;
    var key = '';
    var i = j = 0;
    var len = len2 = 0;
    var strCheck = '0123456789';
    var aux = aux2 = '';
    var whichCode = (window.Event) ? e.which : e.keyCode;
    if (whichCode == 13) return true;
    key = String.fromCharCode(whichCode); // Valor para o código da Chave
    if (strCheck.indexOf(key) == -1) return false; // Chave inválida
    len = objTextBox.value.length;
    for(i = 0; i < len; i++)
        if ((objTextBox.value.charAt(i) != '0') && (objTextBox.value.charAt(i) != SeparadorDecimal)) break;
    aux = '';
    for(; i < len; i++)
        if (strCheck.indexOf(objTextBox.value.charAt(i))!=-1) aux += objTextBox.value.charAt(i);
    aux += key;
    len = aux.length;
    if (len == 0) objTextBox.value = '';
    if (len == 1) objTextBox.value = '0'+ SeparadorDecimal + '0' + aux;
    if (len == 2) objTextBox.value = '0'+ SeparadorDecimal + aux;
    if (len > 2) {
        aux2 = '';
        for (j = 0, i = len - 3; i >= 0; i--) {
            if (j == 3) {
                aux2 += SeparadorMilesimo;
                j = 0;
            }
            aux2 += aux.charAt(i);
            j++;
        }
        objTextBox.value = '';
        len2 = aux2.length;
        for (i = len2 - 1; i >= 0; i--)
        objTextBox.value += aux2.charAt(i);
        objTextBox.value += SeparadorDecimal + aux.substr(len - 2, len);
    }
    return false;
}

 // -->
</script>
<script language="JavaScript" type="text/javascript">
<!--
function FormOk(formulario) {
   var mensagem = '';
   var primeiro = 0;
   for (var i = 0; i < formulario.length; i++) {
       if (formulario.elements[i].getAttribute('erro')) {
          if ((formulario.elements[i].value == '') ||
             (formulario.elements[i].value.length == 0) ||
             (formulario.elements[i].selectedIndex <= 0)
             ) {
             mensagem = mensagem + formulario.elements[i].getAttribute('erro') + "\n";
             // troca a cor da borda do campo
             formulario.elements[i].style.border = '2px solid #03C';
             if (primeiro == 0) { var primeiro = i; }
          }
      }
   }
   if (mensagem != '') {
      formulario.elements[primeiro].focus();
      alert(mensagem);
      return false;
   } else {
       return true;
   }
} 

function checa(nform) {
	//validacao de radio buttons sem saber quantos sao
	marcado = -1
	for (i=0; i<nform.resp.length; i++) {
		if (nform.resp[i].checked) {
			marcado = i
			resposta = nform.resp[i].value
		}
	}
	
	if (marcado == -1) {
		alert("Selecione uma Forma de Pagamento");
		nform.resp[0].focus();
		return false;
	} else { //esse else so foi colocado para evitar que o form desse o submit
		document.form1.submit();
		return false; 
	} 
		return true;
}

//-->
</script>
<script language="javascript">
function alta(valor)
{
valor.value=valor.value.toUpperCase();
}
function pophermes(){
	document.getElementById('pophermes').style.display='block'
	document.getElementById('bodypage').style.display='none'
	}
function popbella(){
	document.getElementById('popbella').style.display='block'
	document.getElementById('bodypage').style.display='none'
	}
function poppvpv(){
	document.getElementById('poppvpv').style.display='block'
	document.getElementById('bodypage').style.display='none'
	}
</script>
</head>
<body>
<%if Application("hlpage") = "inicio" then%>
<div id="pophermes">
	<div id="pop">
    	<div style="position:absolute;z-index:99999;visibility:visible;left:800px;top:-15px;"><a href="./?hl=inicio"><img src="imagens/fechar.png" border="0"></a></div>
		<iframe id="pageframe" style="position:absolute;width:800px;height:580px;background:#FFF;margin:10px;border:0;overflow:hidden;" frameborder="0" scrolling="no" src="catalogo/hermes/"></iframe>
    </div>
</div>
<div id="popbella">
	<div id="pop">
    	<div style="position:absolute;z-index:99999;visibility:visible;left:800px;top:-15px;"><a href="./?hl=inicio"><img src="imagens/fechar.png" border="0"></a></div>
		<iframe id="pageframe" style="position:absolute;width:800px;height:580px;background:#FFF;margin:10px;border:0;overflow:hidden;" frameborder="0" scrolling="no" src="catalogo/bella/"></iframe>
        </div>
    </div>
</div>
<div id="poppvpv">
	<div id="pop">
    	<div style="position:absolute;z-index:99999;visibility:visible;left:800px;top:-15px;"><a href="./?hl=inicio"><img src="imagens/fechar.png" border="0"></a></div>
		<iframe id="pageframe" style="position:absolute;width:800px;height:580px;background:#FFF;margin:10px;border:0;overflow:hidden;" frameborder="0" scrolling="no" src="catalogo/pvpv/"></iframe>
    </div>
</div>
<%end if%>
<div id="bodypage" class="bodypage">
		<div class="header">
			<div class="head">
				<div class="logo"><a href="./"></a>
				</div>
				<div class="menusuperior">
					<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" style="display:block" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="780" height="125">
            		<param name="movie" value="swf/menu.swf">
            		<param name="quality" value="high">
	        		<param name="menu" value="false">
	        		<param name="wmode" value="transparent">
          			<embed wmode="transparent" src="swf/menu.swf" menu="false" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="780" height="125"></embed></object>
				</div>
			</div>
		</div>
		<%
		if not request.querystring("hl") <> "inicio" or not request.querystring("hl") <> "" then
		%>
		<div class="banner">
			<div id="slider">
					<a id="slider-prev">&nbsp;</a>
					<div class="slider-content">
						<ul>
						    <li><img src="imagens/banner/banner1.jpg" alt="" /></li>
						    <li><img src="imagens/banner/banner2.jpg" alt="" /></li>
						    <li><img src="imagens/banner/banner3.jpg" alt="" /></li>
						</ul>
					</div>		
					<a id="slider-next">&nbsp;</a>
					<div class="slider-pagination">
						<ul>
						    <li><a href="#">1</a></li>
						    <li><a href="#">2</a></li>
						    <li><a href="#">3</a></li>
						</ul>
					</div>
					<div class="tl">&nbsp;</div>
					<div class="tr">&nbsp;</div>
					<div class="bl">&nbsp;</div>
					<div class="br">&nbsp;</div>
			</div>
		</div>
		<%
		end if
		%>
		<div class="linetop">
		</div>
		<div class="middle">
			<div class="body">
				<!-- Tamanho da página está definido na class "bodycenter" -->
				<%
				select case request.querystring("hl")
				case ""
				bch = 665
				case "inicio"
				bch = 665
				case "cadastre-se"
				bch = 1450
				case "cadastrar.ok"
				bch = 665
				case "empresa"
				bch = 670
				case "passarpedido"
				bch = 665
				end select
				%>
				<div class="bodycenter" style="height:<%=bch%>px;">
					<div class="menu">
						<ul>
							<li><a href="./?hl=comofunciona">Como Funcionamos</a></li>
							<li><a href="./?hl=cadastre-se">Cadastre-se</a></li>
							<li><a href="./?hl=passarpedido">Passe Seu Pedido</a></li>
							<li><a href="./?hl=pecacatalogos">Peça Catálogo Aqui</a></li>
							<li><a href="./?hl=datacampanha">Data das Campanhas</a></li>
							<li><a href="./?hl=clubedascampeas">Clube das Campeãs</a></li>
							<li><a href="./?hl=clubeamizade">Clube da Amizade</a></li>
							<li><a href="./?hl=pontovaipontovem">Ponto vai Ponto Vem</a></li>
							<li><a href="./?hl=felicidadepremiada">Fidelidade Premiada</a></li>
						</ul>
					</div>
					<div class="corpo">
					<!-- Início das páginas -->
					<%
					if not request.querystring("hl") <> "inicio" or not request.querystring("hl") <> "" then
					if trim(Request.ServerVariables("server_name")) <> "catalogoshermes.com.br" or trim(Request.ServerVariables("server_port")) <> 80 then
					For Each objCookie In Request.Cookies
					  Response.Cookies(objCookie) = ""
					Next
					session.abandon
					if trim(Request.ServerVariables("Query_String")) <> "" then
					vars = "?"&trim(Request.ServerVariables("Query_String"))
					end if
					response.redirect("http://catalogoshermes.com.br/v2/"&vars&"")
					end if
					%>
						<div class="bar"><div style="margin-left:10px;">Confira os Catálogos On-line</div>
						</div>
						<div class="cat">
							<div class="catalogoHermes left" style="width:133px;height:202px;float:left;z-index:9999;">
								<a href="javascript:pophermes();"><img src="catalogo/hermes/paginas/001.jpg" width="123px" height="192px" style="margin:5px;" title="Clique para ver o catálogo" alt="Clique para ver o catálogo" border="0"></a>
							</div>
							<div style="width:172px;height:182px;padding:5px;float:left;">
								<b>Catálogo Hermes<br>
								<br>
								<font size="1">
								<span style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; display: inline; float: none">
								O</span></font><span style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; display: inline !important; float: none"><font size="1"> 
								maior e melhor catálogo de variedades do Brasil.
								</font></span></b>
								<span style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; display: inline !important; float: none; font-weight: 700">
								<font size="1">Mais de 10 Mil itens diferentes</font></span><span style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; display: inline; float: none"><font size="1"><br style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; font-weight: 700; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; background-color: rgb(188, 35, 32)">
								</font>
								<span style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; display: inline !important; float: none; font-weight: 700">
								<font size="1">Cerca de 1,3 Milhões de 
								exemplares.</font></span><font size="1"><br style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; font-weight: 700; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; background-color: rgb(188, 35, 32)">
								</font>
								<span style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; display: inline !important; float: none; font-weight: 700">
								<font size="1">Uma ampla linha de produtos: moda 
								feminina, masculina e infantil, lingerie, 
								artigos de cama, mesa e banho, utilidades 
								domésticas, cosméticos, calçados femininos e 
								masculinos, livros, DVDs e vários outros.</font></span></span></div>
							<div style="width:133px;height:202px;float:left;">
								<a href="javascript:popbella();">
								<img src="catalogo/bella/paginas/001.jpg" width="123px" height="192px" style="margin:5px;" title="Clique para ver o catálogo" alt="Clique para ver o catálogo" border="0"></a>
							</div>
							<div style="width:172px;height:182px;padding:5px;float:left;">
								<b>Catálogo Bella</b><br>
								<br>
								<b>
								<span style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; display: inline !important; float: none">
								<font size="1">O Catálogo de Cosméticos e 
								Bijuterias da Hermes. Quase um milhão de 
								exemplares todo mês.</font></span><font size="1"><br style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px">
								</font>
								<span style="color: rgb(247, 247, 248); font-family: Arial, Verdana; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: justify; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px; display: inline !important; float: none">
								<font size="1">Uma ampla linha de produtos: 
								maquiagem, cremes para o corpo, mãos e pés, 
								cremes antiidades, tratamento para cabelos, 
								shampoos e condicionadores, fragrâncias, 
								sabonetes e vários outros.</font></span></b></div>
						</div>
						<div class="pvpv">
								<a href="javascript:poppvpv();">
								<img src="catalogo/pvpv/paginas/001.jpg" width="123px" height="192px" style="margin:5px;" title="Clique para ver a revista" alt="Clique para ver a revista" border="0"></a>
						</div>					
						<div class="bar2">
							<div style="position:absolute;width:400px;height:200px;margin:10px;"><img src="imagens/sr001.png" width="400px" height="200px"></div>
							<iframe style="position:absolute;margin:10px;left:410px;width:350px;height:200px;" src="http://www.youtube-nocookie.com/embed/qed2UEWop54?wmode=transparent&rel=0" frameborder="0" allowfullscreen></iframe>
						</div>
						<div class="fpag">
							<div style="position:relative;float:left;width:350px;height:130px;text-align:10px;padding:10px;">
							<p style="text-align: justify">
							<font color="#F7CA01"><b>Olha que legal!!!<br><br>
							</b></font><font style="font-size: 9pt"><b>Formas de Pagamento<br>
							</b>A consultora Hermes terá um prazo de 12 à 14 dias 
							para pagar o boleto (CREDHERMES), após a entrega da caixa. Ou ainda, 
							a consultora Hermes ou a cliente poderão pagar 
							parceladamente em até 10x sem juros, no seu cartão 
							de crédito.</font></div>
						</div>
					<%end if%>
					<%
					if not request.querystring("hl") <> "empresa" then
					if trim(Request.ServerVariables("server_name")) <> "catalogoshermes.com.br" or trim(Request.ServerVariables("server_port")) <> 80 then
					For Each objCookie In Request.Cookies
					  Response.Cookies(objCookie) = ""
					Next
					session.abandon
					if trim(Request.ServerVariables("Query_String")) <> "" then
					vars = "?"&trim(Request.ServerVariables("Query_String"))
					end if
					response.redirect("http://catalogoshermes.com.br/v2/"&vars&"")
					end if
					%>
                                  <div id="text_int">
                <img border="0" src="imagens/hermes.png" align="left" alt="Sede da Hermes S/A" title="Sede da Hermes S/A"><h1 align="justify">
				A empresa Hermes</h1>
                <p align="justify">Fundada em 1942, e sediada na cidade do Rio 
				de Janeiro, a Hermes S/A, é a maior empresa brasileira de vendas 
				por catálogos de variedades, com mais de 10 milhões de clientes 
				em todo o Brasil. Possui uma grande estrutura de Distribuição e 
				Atendimento, com 32.000m² de área construída e, capacidade para 
				atender cerca de 40 mil pedidos por dia.<br>
                <br>
              	Atua através de diversos canais de vendas:<br>
              	Por meio dos catálogos Hermes e Bella;<br>
              	Pelo site: <a href="#">www.catalogoshermes.com.br</a><br><br>
                <br>
              	Atualmente a empresa oferece uma grande variedade de produtos, 
				como utilidades domésticas, confecções masculina,&nbsp; feminina 
				e infantil, linha íntima, calçados, artigos de cama, mesa e 
				banho, cosméticos, eletônicos, eletrodomésticos, 
				eletroportáteis, perfumaria, telefonia e celulares, cine e foto 
				e vários outros. São mais de 30 mil itens comercializados 
				diariamente. Com a missão de proporcionar comodidade e bom 
				atendimento, aliados ao crescimento, desenvolvimento e 
				produtividade, atua com responsabiliddde social, e possui mais 
				de 1500 funcionários treinados e especializados.<br /></p><br />
                <h1 align="justify">Catálogos Hermes</h1>
                <p align="justify">Descrição</p>
                <p align="justify">&nbsp;</p></div>
                    <%end if%>
					<%
					if not request.querystring("hl") <> "cadastre-se" then
					if trim(Request.ServerVariables("server_name")) <> "ssl6260.websiteseguro.com" or trim(Request.ServerVariables("server_port")) <> 443 then
					For Each objCookie In Request.Cookies
					  Response.Cookies(objCookie) = ""
					Next
					session.abandon
					if trim(Request.ServerVariables("Query_String")) <> "" then
					vars = "?"&trim(Request.ServerVariables("Query_String"))
					end if
					response.redirect("https://ssl6260.websiteseguro.com/catalogoshermes1/v2/"&vars&"")
					end if
					%>
						<div class="bar"><div style="margin-left:10px;">Cadastre-se</div>
						</div>
						<div id="text_int" style="margin-top:10px;margin-bottom:10px;background:#A81716;">
                Preencha o formulário abaixo, para 
				você se cadastrar e começar a revender os produtos HERMES e 
				BELLA. 
                Você também poderá enviar o formulário impresso entregando 
				diretamente em nossa loja. Para imprimir a ficha de cadastro,<b>
                <a href="?"><font color="#FFFFFF">
				<span style="text-decoration: none">clique aqui</span></font></a></b>.</div>
                <form action="?">
                <input type="hidden" name="action" value="cadastrar" size="20">
                <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="768">
                  <tr>
                    <td width="768" height="30" colspan="3">
                    <div id="text_int"><b><font color="#F8D000">Obs.:</font></b><br>
                      	- Todos os campos marcados com <font color="#F8D000">*</font> são obrigatórios.<br>
                      	- Para efetuar o preenchimento dos campos, CEP, 
						DDD+Telefone, DDD+Celular, CPF, Identidade, digite somente os 
						números, sem pontuação ou espaços.<i><br>
&nbsp;</i></div><br /></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Nome 
						Completo:
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
              <input name="nome" type="text" class="input1" size="42" erro="Seu Nome" onkeyup="javascript:alta(this);" value="<%=nome%>">
              <font color="#F8D000" size="2"><b><%="  "&session("erro_nome")%></b></font></td>
                  </tr>
                  <tr>
                    <td width="192" height="40">
                    <div id="text_int">Endereço: <b><font color="#F8D000">*</font></b></div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
              <input name="endereco" type="text" class="input1" size="42" erro="Seu Endereço" onkeyup="javascript:alta(this);" value="<%=endereco%>">
              <font color="#F8D000" size="2"><b><%="  "&session("erro_endereco")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Ref. para 
						Entrega:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left"> 
              <input name="referencia" type="text" class="form" id="referencia" size="32" onkeyup="javascript:alta(this);" value="<%=referencia%>"></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Bairro: <b>
                      <font color="#F8D000">*</font></b></div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left"> 
              <input name="bairro" type="text" class="form" id="bairro" size="32" erro="Seu Bairro" onkeyup="javascript:alta(this);" value="<%=bairro%>">
              <font color="#F8D000" size="2"><b><%="  "&session("erro_bairro")%></b></font>
					</td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">CEP: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
              <input name="cep" type="text" class="form" id="cep" size="10" maxlength="8" erro="Seu CEP" onkeypress="return CEP(this, event)" value="<%=cep%>">
              <font color="#F8D000" size="2"><b><%="  "&session("erro_cep")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Cidade: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
              <input name="cidade" type="text" class="form" id="cidade" size="32" erro="Sua Cidade" onkeyup="javascript:alta(this);" value="<%=cidade%>">
              <font color="#F8D000" size="2"><b><%="  "&session("erro_cidade")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Estado:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
<select name="estado" class="form" id="estado">
  
<option value="ACRE" <%if estado = "ACRE" then%>selected="selected"<%end if%>>
ACRE
</option><option value="ALAGOAS" <%if estado = "ALAGOAS" then%>selected="selected"<%end if%>>
ALAGOAS
</option><option value="AMAPÁ" <%if estado = "AMAPÁ" then%>selected="selected"<%end if%>>
AMAPÁ
</option><option value="AMAZONAS" <%if estado = "AMAZONAS" then%>selected="selected"<%end if%>>
AMAZONAS
</option><option value="BAHIA" <%if estado = "BAHIA" then%>selected="selected"<%end if%>>
BAHIA
</option><option value="CEARÁ" <%if estado = "CEARÁ" then%>selected="selected"<%end if%>>
CEARÁ
</option><option value="DISTRITO FEDERAL" <%if estado = "DISTRITO FEDERAL" then%>selected="selected"<%end if%>>
DISTRITO FEDERAL
</option><option value="ESPÍRITO SANTO" <%if estado = "ESPÍRITO SANTO" then%>selected="selected"<%end if%>>
ESPÍRITO SANTO
</option><option value="GOIÁS" <%if estado = "GOIÁS" then%>selected="selected"<%end if%>>
GOIÁS
</option><option value="MARANHÃO" <%if estado = "MARANHÃO" then%>selected="selected"<%end if%>>
MARANHÃO
</option><option value="MATO GROSSO" <%if estado = "MATO GROSSO" then%>selected="selected"<%end if%>>
MATO GROSSO
</option><option value="MATO GROSSO DO SUL" <%if estado = "MATO GROSSO DO SUL" then%>selected="selected"<%end if%>>
MATO GROSSO DO SUL
</option><option value="MINAS GERAIS" <%if estado = "MINAS GERAIS" then%>selected="selected"<%end if%>>
MINAS GERAIS
</option><option value="PARÁ" <%if estado = "PARÁ" then%>selected="selected"<%end if%>>
PARÁ
</option><option value="PARAÍBA" <%if estado = "PARAÍBA" then%>selected="selected"<%end if%>>
PARAÍBA
</option><option value="PARANÁ" <%if estado = "PARANÁ" then%>selected="selected"<%end if%>>
PARANÁ
</option><option value="PERNAMBUCO" <%if estado = "PERNAMBUCO" then%>selected="selected"<%end if%>>
PERNAMBUCO
</option><option value="PIAUÍ" <%if estado = "PIAUÍ" then%>selected="selected"<%end if%>>
PIAUÍ
</option><option value="RIO DE JANEIRO" <%if estado = "RIO DE JANEIRO" or estado = "" then%>selected="selected"<%end if%>>
RIO DE JANEIRO
</option><option value="RIO GRANDE DO NORTE" <%if estado = "RIO GRANDE DO NORTE" then%>selected="selected"<%end if%>>
RIO GRANDE DO NORTE
</option><option value="RIO GRANDE DO SUL" <%if estado = "RIO GRANDE DO SUL" then%>selected="selected"<%end if%>>
RIO GRANDE DO SUL
</option><option value="RONDÔNIA" <%if estado = "RONDÔNIA" then%>selected="selected"<%end if%>>
RONDÔNIA
</option><option value="RORAIMA" <%if estado = "RORAIMA" then%>selected="selected"<%end if%>>
RORAIMA
</option><option value="SANTA CATARINA" <%if estado = "SANTA CATARINA" then%>selected="selected"<%end if%>>
SANTA CATARINA
</option><option value="SÃO PAULO" <%if estado = "SÃO PAULO" then%>selected="selected"<%end if%>>
SÃO PAULO
</option><option value="SERGIPE" SERGIPE>SERGIPE
</option><option value="TOCANTINS" <%if estado = "TOCANTINS" then%>selected="selected"<%end if%>>
TOCANTINS
		  </option></select></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Sexo:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
              <select name="sexo" class="form" id="sexo">
                <option value="FEMININO" <%if sexo = "FEMININO" then%>selected="selected"<%end if%>>
				FEMININO</option>
                <option value="MASCULINO" <%if sexo = "MASCULINO" then%>selected="selected"<%end if%>>
				MASCULINO</option>
              </select></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">DDD+Telefone: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
              <input name="tel" type="text" class="form" id="tel" size="14" maxlength="10" value="<%=tel%>">
              <font color="#F8D000" size="2"><b><%="  "&session("erro_tel")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">DDD+Celular:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
<input name="cel" type="text" class="form" id="cel" size="14" erro="Seu Telefone" maxlength="10" value="<%=cel%>"></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Email:<font color="#F8D000"><b> 
						*</b></font></div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
              <input name="email" type="text" class="form" id="email" size="35" value="<%=email%>">
              <font color="#F8D000" size="2"><b><%="  "&session("erro_email")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Estado Civíl:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
              
              <select name="civil" class="form" id="civil">
                <option value="2 - SOLTEIRO(A)" <%if civil = "2 - SOLTEIRO(A)" then%>selected="selected"<%end if%>>
				SOLTEIRO(A)</option>
                <option value="1 - CASADO(A)" <%if civil = "1 - CASADO(A)" then%>selected="selected"<%end if%>>
				CASADO(A)</option>
                <option value="4 - DESQUITADO(A)/DIVORCIADO(A)" <%if civil = "4 - DESQUITADO(A)/DIVORCIADO(A)" then%>selected="selected"<%end if%>>
				DESQUITADO(A)/DIVORCIADO(A)</option>
                <option value="3 - VIÚVO(A)" <%if civil = "3 - VIÚVO(A)" then%>selected="selected"<%end if%>>
				VIÚVO(A)</option>
              </select></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Data de 
						Nascimento:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
<input name="nascimento" type="text" class="form" id="nascimento" onkeyup="javascript:alta(this);" size="10" maxlength="10" onkeypress="return MaskData(this, event)" value="<%=nascimento%>"></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">CPF:<font color="#F8D000"><b> 
						*</b></font></div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
<input name="cpf" type="text" class="form" id="cpf" size="14" maxlength="11" value="<%=cpf%>">
<font color="#F8D000" size="2"><b><%="  "&session("erro_cpf")%></b></font>
</td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Identidade 
						(RG): 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left"> 
<input name="id" type="text" class="form" id="id" size="10" maxlength="10" value="<%=id%>">
<font color="#F8D000" size="2"><b><%="  "&session("erro_id")%></b></font>
</td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Órgão Exp.:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
<input name="orgao" type="text" class="form" id="orgao" onkeyup="javascript:alta(this);" size="10" maxlength="10" erro="Seu Órgão Expedidor" value="<%=orgao%>"></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Tempo de 
						Residência:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left"> 
<input name="tresidencia" type="text" class="form" id="tresidencia" size="14" erro="Seu Tempo de Residência" onkeyup="javascript:alta(this);" value="<%=tresidencia%>"></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Situação da 
						Residência:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
<select name="sresidencia" class="form" id="sresidencia">
  <option value="1 - PRÓPRIA" <%if sresidencia = "1 - PRÓPRIA" then%>selected="selected"<%end if%>>
	PRÓPRIA</option>
  <option value="2 - ALUGADA" <%if sresidencia = "2 - ALUGADA" then%>selected="selected"<%end if%>>
	ALUGADA</option>
  <option value="3 - PAIS" <%if sresidencia = "3 - PAIS" then%>selected="selected"<%end if%>>
	PAIS</option>
  <option value="4 - PARENTES" <%if sresidencia = "4 - PARENTES" then%>selected="selected"<%end if%>>
	PARENTES</option>
  <option value="5 - OUTROS" <%if sresidencia = "5 - OUTROS" then%>selected="selected"<%end if%>>
	OUTROS</option>
</select></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Nome da Mãe:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left"> 
<input name="nmae" type="text" class="form" id="nmae" size="42" onkeyup="javascript:alta(this);" value="<%=nmae%>"></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Data de 
						Nascimento da Mãe:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left"> 
<input name="dmae" type="text" class="form" id="dmae" onkeyup="javascript:alta(this);" size="10" maxlength="10" onkeypress="return MaskData(this, event)" value="<%=dmae%>"></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Nome do Pai:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
<input name="npai" type="text" class="form" id="npai" size="42" onkeyup="javascript:alta(this);" value="<%=npai%>"></td>
                  </tr>
                  <tr>
                    <td width="192" height="40"><div id="text_int">Data de 
						Nascimento do Pai:</div></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40" align="left">
<input name="dpai" type="text" class="form" id="dpai" onkeyup="javascript:alta(this);" size="10" maxlength="10" onkeypress="return MaskData(this, event)" value="<%=dpai%>"></td>
                  </tr>
                  <tr>
                    <td width="768" height="40" colspan="3" align="left"><div id="text_int">
						Nome e Data de Nascimento dos Filhos Menores de 14 anos:</div></td>
                  </tr>
                  <tr>
                    <td width="768" height="40" colspan="3" align="left">
<input name="nfilho1" type="text" class="form" id="nfilho1" size="40" onkeyup="javascript:alta(this);" value="<%=nfilho1%>">
<input name="dfilho1" type="text" class="form" id="dfilho1" onkeyup="javascript:alta(this);" size="10" maxlength="10" onkeypress="return MaskData(this, event)" value="<%=dfilho1%>"></td>
                  </tr>
                  <tr>
                    <td width="768" height="40" colspan="3" align="left">
<input name="nfilho2" type="text" class="form" id="nfilho2" size="40" onkeyup="javascript:alta(this);" value="<%=nfilho2%>">
<input name="dfilho2" type="text" class="form" id="dfilho2" onkeyup="javascript:alta(this);" size="10" maxlength="10" onkeypress="return MaskData(this, event)" value="<%=dfilho2%>"></td>
                  </tr>
                  <tr>
                    <td width="768" height="40" colspan="3" align="left">
<input name="nfilho3" type="text" class="form" id="nfilho3" size="40" onkeyup="javascript:alta(this);" value="<%=nfilho3%>">
<input name="dfilho3" type="text" class="form" id="dfilho3" onkeyup="javascript:alta(this);" size="10" maxlength="10" onkeypress="return MaskData(this, event)" value="<%=dfilho3%>"></td>
                  </tr>
                  <tr>
                    <td width="768" height="40" colspan="3" align="left">
<input name="nfilho4" type="text" class="form" id="nfilho4" size="40" onkeyup="javascript:alta(this);" value="<%=nfilho4%>">
<input name="dfilho4" type="text" class="form" id="dfilho4" onkeyup="javascript:alta(this);" size="10" maxlength="10" onkeypress="return MaskData(this, event)" value="<%=dfilho4%>"></td>
                  </tr>
                  <tr>
                    <td width="768" height="40" colspan="3" align="left"><div id="text_int">
						<font style="font-size: 9pt">Revendo 
						Natura, Avon e outros: </font>
<input name="revendooutros" type="checkbox" id="revendooutros" value="Sim" <%if revendooutros = "Sim" then%>checked<%end if%>></div></td>
                  </tr>
                  <tr>
                    <td width="192" height="40" align="left">
              <input name="Reset" type="reset" class="botaoform" id="button" value=" Limpar ">&nbsp;&nbsp;<input name="button2" type="submit" class="botaoform" id="button2" value=" Cadastrar "></td>
                    <td width="5" height="40">&nbsp;</td>
                    <td width="570" height="40">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="192" height="30">&nbsp;</td>
                    <td width="5" height="30">&nbsp;</td>
                    <td width="570" height="30">&nbsp;</td>
                  </tr>
                </table>
                </form>
					<%
					end if

					if not request.querystring("hl") <> "cadastrar.ok" then
					if trim(Request.ServerVariables("server_name")) <> "ssl6260.websiteseguro.com" or trim(Request.ServerVariables("server_port")) <> 443 then
					For Each objCookie In Request.Cookies
					  Response.Cookies(objCookie) = ""
					Next
					session.abandon
					if trim(Request.ServerVariables("Query_String")) <> "" then
					vars = "?"&trim(Request.ServerVariables("Query_String"))
					end if
					response.redirect("https://ssl6260.websiteseguro.com/catalogoshermes1/v2/"&vars&"")
					end if%>
                    <br>
              <br>
              <br>
              <br>
              <br>
              <br>
              <font color="#F8D000"><b>Cadastro realizado com sucesso!</b></font>
					<%
					end if
					
					if not request.querystring("hl") <> "passarpedido" then
					if trim(Request.ServerVariables("server_name")) <> "ssl6260.websiteseguro.com" or trim(Request.ServerVariables("server_port")) <> 443 then
					For Each objCookie In Request.Cookies
					  Response.Cookies(objCookie) = ""
					Next
					session.abandon
					if trim(Request.ServerVariables("Query_String")) <> "" then
					vars = "?"&trim(Request.ServerVariables("Query_String"))
					end if
					response.redirect("https://ssl6260.websiteseguro.com/catalogoshermes1/v2/"&vars&"")
					end if
					%>
<div id="text_int"><h1>
                Faça o seu pedido</h1>
                Através do formulário abaixo você poderá enviar o seu pedido. 
				Basta preenchê-lo corretamente, você terá todas as informações 
				necessárias para efetuar os seus pedidos de forma mais rápida.<br>
                <br>
                Após preencher os seus pedidos, clique em <b>INFORMAÇÕES DE 
				PAGAMENTO</b>.<br>
                <br>
                Preencha os campos: <b>Nome, E-mail, Endereço, Bairro, Cidade, 
				Telefone com um número atualizado e o seu Código de consultora 
				se tiver</b>. Depois Preencha as opções de Pagamento e Enviar.<br>
&nbsp;</div>
					<%
					end if
					%>                    
                    <!-- Fim das páginas -->
					</div>
				</div>
			</div>
		</div>
		<div class="linebottom">
		</div>
		<div class="rodape">
			<div class="ralign">
			<div class="dev">
				<font color="#8A0608" style="font-family: Arial, Verdana; font-size: 13px; font-style: normal; font-variant: normal; font-weight: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px;">
				<b><br>
				<br>
				Nossa Loja</b><br>
				Fica atrás da Casa &amp; Vídeo, na galeria ao lado - Primeira loja.<br>
				Rua: Gil de Góis, 103 loja n.o 34 - Centro - Campos dos Goytacazes - Rio 
				de Janeiro.</font><font color="#A02C27" style="font-family: Arial, Verdana; font-size: 13px; font-style: normal; font-variant: normal; font-weight: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px;"><br>
				<br>
				</font>
				<b style="color: rgb(102, 102, 102); font-family: Arial, Verdana; font-size: 13px; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: normal; orphans: 2; text-align: -webkit-center; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-size-adjust: auto; -webkit-text-stroke-width: 0px;">
				<font color="#8A0608">Copyright&nbsp;<%=year(now)%> © Catálogo Hermes. Todos os 
				direitos reservados.</font></b>
				<div class="ssl"><%=Application("chacessos")%></div>
			</div>
			</div>
		</div>
<div id="redessociais" style="position:absolute;left:50%;top:100px;visibility:visible;">
		<a href="http://www.facebook.com/catalogos.hermes.9" target="ch_social"><img name="facebook" id="facebook" alt="Facebook" title="Facebook" src="imagens/facebook.png" style="position:absolute;left:425px" border="0"></a>
		<img name="twitter" id="twitter" alt="Twitter" title="Twitter" src="imagens/twitter.png" style="position:absolute;left:460px;" border="0">
</div>
</div>
</body>
</html>
