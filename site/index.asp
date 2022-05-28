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
StrTopo = StrTopo & "                  Fone: (22) 30152376 ou TIM (22) 998796765" & vbNewline
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
oMail.UserName = "admin@catalogoshermes.com.br"
oMail.PassWord = "art@hermes123"
oMail.From = "admin@catalogoshermes.com.br"
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
oMail.UserName = "admin@catalogoshermes.com.br"
oMail.PassWord = "art@hermes123"
oMail.From = "admin@catalogoshermes.com.br"
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
oMail.UserName = "admin@catalogoshermes.com.br"
oMail.PassWord = "art@hermes123"
oMail.From = "admin@catalogoshermes.com.br"
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
oMail.UserName = "admin@catalogoshermes.com.br"
oMail.PassWord = "art@hermes123"
oMail.From = "admin@catalogoshermes.com.br"
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
<head>
<title>..:: Catálogos Hermes ::..</title>
<link rel="shortcut icon" type="image/vnd.microsoft.icon" href="image/flavicon.png" />
<meta http-equiv="X-UA-Compatible" content="IE=9" />
<link rel="stylesheet" type="text/css" href="style/styles.css"/>
<link href="style/SpryTabbedPanels.css" rel="stylesheet" type="text/css">
<link href="style/hermes.css" rel="stylesheet" type="text/css">
<script src="js/SpryTabbedPanels.js" type="text/javascript"></script>
<script type="text/javascript" src="js/jquery.js"></script>
<%if Application("hlpage") = "inicio" then%>
<script src="js/jquery-1.4.2.js" type="text/javascript"></script>
<script src="js/jquery.jcarousel.js" type="text/javascript"></script>
<script src="js/functions.js" type="text/javascript"></script>
<%end if%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="google-site-verification" content="JLljOIf1PqEQ9zYlZlzpSxTMxyRAvAjOtc_VfZnnV2k" />
<script>
 $(document).ready(function(){
  $("dd:not(:first)").hide();
  $("dt a").click(function(){
   $("dd:visible").slideUp("slow");
   $(this).parent().next().slideDown("slow");
   return false;
  });
 });
 </script>
 <script type="text/javascript" language="JavaScript">
<!--
function add_bookmark() {
var browsName = navigator.appName;
if (browsName == "Microsoft Internet Explorer") {
window.external.AddFavorite('http://www.catalogoshermes.com.br','Catálogos Hermes' );
} else if (browsName == "Netscape") {
alert ("Para adicionar nosso site aos seus Favoritos aperte CTRL+D");
}
}
// -->
</script>
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
</script>
<script language="javascript"> 
function catalogo (URL){
   window.open(URL,"catalogo","width=800,height=580,scrollbars=NO")
} 
</script>
<%
if Application("hlpage") = "passarpedido" then
%>
<meta http-equiv="X-UA-Compatible" content="IE=9" />
<%
end if
%>
</head>

<body topmargin="0" leftmargin="0">
<div id="redessociais" style="position: absolute; left: 50%;">
<a href="http://www.facebook.com/catalogos.hermes.9" target="ch_social">
<img name="facebook" id="facebook" alt="Facebook" title="Facebook" src="image/facebook.png" class="imgface" border="0"></a>
<img name="twitter" id="twitter" alt="Twitter" title="Twitter" src="image/twitter.png" class="imgtweet" border="0">
</div>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="100%">
  <tr>
    <td width="100%" height="125" align="center">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="988" height="125">
      <tr>
        <td width="208" align="center" valign="middle"><a href="?hl=inicio">
        <img border="0" src="image/logo.png"></a></td>
        <td width="780" align="right" valign="top" background="image/cmdb.png">
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" style="display:block" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="780" height="125">
            <param name="movie" value="swf/menu.swf">
            <param name="quality" value="high">
	        <param name="menu" value="false">
	        <param name="wmode" value="transparent">
          <embed wmode="transparent" src="swf/menu.swf" menu="false" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="780" height="125"></embed></object>
        </td>
      </tr>
    </table>
    <%if Application("hlpage") = "inicio" then%>
    <div class="banner">
		<div id="slider">
				<a id="slider-prev">&nbsp;</a>
				<div class="slider-content">
					<ul>
					    <li><img src="image/banner/banner1.jpg" alt="" /></li>
					</ul>
				</div>		
				<a id="slider-next">&nbsp;</a>
				<div class="slider-pagination">
					<ul>
					    <li><a href="#">1</a></li>
					</ul>
				</div>
				<div class="tl">&nbsp;</div>
				<div class="tr">&nbsp;</div>
				<div class="bl">&nbsp;</div>
				<div class="br">&nbsp;</div>
		</div>
	</div>
    <%end if%>
    </td>
  </tr>
  <tr>
    <td width="100%" align="center" height="10"><span style="font-size: 1pt">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="100%" align="center" height="14" bgcolor="#8A0608">
    <span style="font-size: 1pt">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="100%" align="center" valign="top" bgcolor="#BC2320">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="988" height="100%">
      <tr>
        <td height="100%" align="center" valign="top">
        <%
        select case Application("hlpage")
        case "inicio"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="988">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>
              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="773">
                  <tr>
                    <td width="773" height="182" align="center" colspan="5">
                    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="773" height="182">
                      <tr>
                        <td width="773" colspan="3"><a href="?hl=premiofantastico">
                        <img border="0" src="image/200000017.jpg" width="773" height="190"></a></td>
                      </tr>
                      <tr>
                        <td width="572">&nbsp;</td>
                        <td width="16">&nbsp;</td>
                        <td width="185">&nbsp;</td>
                      </tr>
                      <tr>
                        <td width="572"><a href="?hl=superindiqueeganhe">
                        <img border="0" src="image/indique.jpg"></a></td>
                        <td width="16">&nbsp;</td>
                        <td width="185"><a href="javascript:catalogo('window/')"><img border="0" src="image/pvpv.jpg"></a></td>
                      </tr>
                    </table>
                    </td>
                  </tr>
                  <tr>
                    <td width="247" align="center">&nbsp;
                    </td>
                    <td width="16">&nbsp;</td>
                    <td width="247" align="center">&nbsp;
                    </td>
                    <td width="16">&nbsp;</td>
                    <td width="247" align="center">&nbsp;
                    </td>
                  </tr>
                  <tr>
                    <td width="247" height="102" align="center">
                    <a href="?hl=cadastre-se">
                    <img border="0" src="image/bg_001.png"></a></td>
                    <td width="16" height="102">&nbsp;</td>
                    <td width="247" height="102" align="center">
                    <a href="?hl=passarpedido">
                    <img border="0" src="image/bg_002.png"></a></td>
                    <td width="16" height="102">&nbsp;</td>
                    <td width="247" height="102" align="center">
                    <a href="javascript:catalogo('windows/')">
                    <img border="0" src="image/bg_003.png"></a></td>
                  </tr>
                  <tr>
                    <td width="713" colspan="5">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="713" colspan="5">
                  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="773">
                  <tr>
                    <td width="379" align="left" valign="top" colspan="3">
                    <div id="text_int"><h1>
                      <img border="0" src="image/capa_hermes.png" align="right" hspace="10">Catálogo 
						Hermes</h1>
                    <p align="justify">O maior e melhor catálogo de variedades 
					do Brasil.<br>
                    Mais de 10 Mil itens diferentes<br>
                    Cerca de 1,3 Milhões de exemplares<br>
                    Uma ampla linha de produtos: moda feminina, masculina e 
					infantil, lingerie, artigos de cama, mesa e banho, 
					utilidades domésticas, cosméticos, calçados femininos e 
					masculinos, livros, DVDs e vários outros.</div></td>
                    <td width="15">&nbsp;</td>
                    <td width="385" align="left" valign="top" colspan="3">
                    <h1>
                    <img border="0" src="image/capa_bella.png" align="right" hspace="10">Catálogo 
					Bella</h1>
                    <div id="text_int"><p align="justify">O Catálogo de 
						Cosméticos e Bijuterias da Hermes.<br>
          				itens diferentes<br>
                    	Quase um milhão de exemplares todo mês.<br>
                    	Uma ampla linha de produtos: maquiagem, cremes para o 
						corpo, mãos e pés, cremes antiidades, tratamento para 
						cabelos, shampoos e condicionadores, fragrâncias, 
						sabonetes e vários outros.</div></td>
                  </tr>
                  <tr>
                    <td width="379" align="left" valign="top" colspan="3">&nbsp;
                    </td>
                    <td width="15">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="779" align="left" valign="top" colspan="7" height="115">
                    <img border="0" src="image/parcelamento.jpg" width="773" height="115"></td>
                  </tr>
                  <tr>
                    <td width="190" align="left" valign="top">&nbsp;
                    </td>
                    <td width="15" align="left" valign="top">&nbsp;
                    </td>
                    <td width="94" align="left" valign="top">&nbsp;
                    </td>
                    <td width="15">&nbsp;</td>
                    <td width="97" align="left" valign="top">&nbsp;
                    </td>
                    <td width="15" align="left" valign="top">&nbsp;
                    </td>
                    <td width="192" align="left" valign="top">&nbsp;
                    </td>
                  </tr>
                  </table>
                    </td>
                  </tr>
                </table>
            </div></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "empresa"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>
                            </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
                <img border="0" src="image/hermes.png" align="left" alt="Sede da Hermes S/A" title="Sede da Hermes S/A"><h1 align="justify">
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
              	Pelo site: <a href="#">www.catalogoshermes.com.br</a><br>
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
                <p align="justify">&nbsp;</p></div></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "cadastre-se"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Cadastre-se</h1>
                <p align="justify"><strong>Preencha o formulário abaixo, para 
				você se cadastrar e começar a revender os produtos HERMES e 
				BELLA.</strong> 
                Você também poderá enviar o formulário impresso entregando 
				diretamente em nossa loja. Para imprimir a ficha de cadastro,
                <a href="?">clique aqui</a>.<br>&nbsp;</p>
                <form name="form" action="?">
                <input type="hidden" name="action" value="cadastrar" size="20">
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="773">
                  <tr>
                    <td width="773" height="30" colspan="3">
                    <div id="text_int"><i><b>Obs.:<br>
                      </b>- Todos os campos marcados com<font color="#F8D000"><b> 
						*
                      </b></font>são obrigatórios.<br>
                      	- Para efetuar o preenchimento dos campos, CEP, 
						Telefone, Celular, CPF, Identidade, digite somente os 
						números, sem pontuação ou espaços.</i></div><br /></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Nome 
						Completo:
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="nome" type="text" class="form" id="nome" size="42" erro="Seu Nome" onKeyUp="javascript:alta(this);" value="<%=nome%>">
              <font color="#F8D000" size="2"><b><%=session("erro_nome")%></b></font></td>
                  </tr>
                  <tr>
                    <td width="184" height="30">
                    <div id="text_int">Endereço: <b><font color="#F8D000">*</font></b></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="endereco" type="text" class="form" id="endereco" size="42" erro="Seu Endereço" onKeyUp="javascript:alta(this);" value="<%=endereco%>">
              <font color="#F8D000" size="2"><b><%=session("erro_endereco")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Ref. para 
						Entrega:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30"> 
              <input name="referencia" type="text" class="form" id="referencia" size="32" onKeyUp="javascript:alta(this);" value="<%=referencia%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Bairro: <b>
                      <font color="#F8D000">*</font></b></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30"> 
              <input name="bairro" type="text" class="form" id="bairro" size="32" erro="Seu Bairro" onKeyUp="javascript:alta(this);" value="<%=bairro%>">
              <font color="#F8D000" size="2"><b><%=session("erro_bairro")%></b></font>
					</td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">CEP: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="cep" type="text" class="form" id="cep" size="10" maxlength="8" erro="Seu CEP" onKeyPress="return CEP(this, event)" value="<%=cep%>">
              <font color="#F8D000" size="2"><b><%=session("erro_cep")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Cidade: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="cidade" type="text" class="form" id="cidade" size="32" erro="Sua Cidade" onKeyUp="javascript:alta(this);" value="<%=cidade%>">
              <font color="#F8D000" size="2"><b><%=session("erro_cidade")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Estado:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
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
                    <td width="184" height="30"><div id="text_int">Sexo:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <select name="sexo" class="form" id="sexo">
                <option value="FEMININO" <%if sexo = "FEMININO" then%>selected="selected"<%end if%>>
				FEMININO</option>
                <option value="MASCULINO" <%if sexo = "MASCULINO" then%>selected="selected"<%end if%>>
				MASCULINO</option>
              </select></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">DDD+Telefone: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="tel" type="text" class="form" id="tel" size="14" maxlength="10" value="<%=tel%>">
              <font color="#F8D000" size="2"><b><%=session("erro_tel")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">DDD+Celular:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
<input name="cel" type="text" class="form" id="cel" size="14" erro="Seu Telefone" maxlength="10" value="<%=cel%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Email:<font color="#F8D000"><b> 
						*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="email" type="text" class="form" id="email" size="35" value="<%=email%>">
              <font color="#F8D000" size="2"><b><%=session("erro_email")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Estado Civíl:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              
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
                    <td width="184" height="30"><div id="text_int">Data de 
						Nascimento:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
<input name="nascimento" type="text" class="form" id="nascimento" onKeyUp="javascript:alta(this);" size="10" maxlength="10" onKeyPress="return MaskData(this, event)" value="<%=nascimento%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">CPF:<font color="#F8D000"><b> 
						*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
<input name="cpf" type="text" class="form" id="cpf" size="14" maxlength="11" value="<%=cpf%>">
<font color="#F8D000" size="2"><b><%=session("erro_cpf")%></b></font>
</td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Identidade 
						(RG): 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30"> 
<input name="id" type="text" class="form" id="id" size="10" maxlength="10" value="<%=id%>">
<font color="#F8D000" size="2"><b><%=session("erro_id")%></b></font>
</td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Órgão Exp.:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
<input name="orgao" type="text" class="form" id="orgao" onKeyUp="javascript:alta(this);" size="10" maxlength="10" erro="Seu Órgão Expedidor" value="<%=orgao%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Tempo de 
						Residência:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30"> 
<input name="tresidencia" type="text" class="form" id="tresidencia" size="14" erro="Seu Tempo de Residência" onKeyUp="javascript:alta(this);" value="<%=tresidencia%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Situação da 
						Residência:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
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
                    <td width="184" height="30"><div id="text_int">Nome da Mãe:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30"> 
<input name="nmae" type="text" class="form" id="nmae" size="42" onKeyUp="javascript:alta(this);" value="<%=nmae%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Data de 
						Nascimento da Mãe:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30"> 
<input name="dmae" type="text" class="form" id="dmae" onKeyUp="javascript:alta(this);" size="10" maxlength="10" onKeyPress="return MaskData(this, event)" value="<%=dmae%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Nome do Pai:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
<input name="npai" type="text" class="form" id="npai" size="42" onKeyUp="javascript:alta(this);" value="<%=npai%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Data de 
						Nascimento do Pai:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
<input name="dpai" type="text" class="form" id="dpai" onKeyUp="javascript:alta(this);" size="10" maxlength="10" onKeyPress="return MaskData(this, event)" value="<%=dpai%>"></td>
                  </tr>
                  <tr>
                    <td width="773" height="30" colspan="3"><div id="text_int">
						Nome e Data de Nascimento dos Filhos Menores de 14 anos:</div></td>
                  </tr>
                  <tr>
                    <td width="773" height="30" colspan="3">
<input name="nfilho1" type="text" class="form" id="nfilho1" size="40" onKeyUp="javascript:alta(this);" value="<%=nfilho1%>">
<input name="dfilho1" type="text" class="form" id="dfilho1" onKeyUp="javascript:alta(this);" size="10" maxlength="10" onKeyPress="return MaskData(this, event)" value="<%=dfilho1%>"></td>
                  </tr>
                  <tr>
                    <td width="773" height="30" colspan="3">
<input name="nfilho2" type="text" class="form" id="nfilho2" size="40" onKeyUp="javascript:alta(this);" value="<%=nfilho2%>">
<input name="dfilho2" type="text" class="form" id="dfilho2" onKeyUp="javascript:alta(this);" size="10" maxlength="10" onKeyPress="return MaskData(this, event)" value="<%=dfilho2%>"></td>
                  </tr>
                  <tr>
                    <td width="773" height="30" colspan="3">
<input name="nfilho3" type="text" class="form" id="nfilho3" size="40" onKeyUp="javascript:alta(this);" value="<%=nfilho3%>">
<input name="dfilho3" type="text" class="form" id="dfilho3" onKeyUp="javascript:alta(this);" size="10" maxlength="10" onKeyPress="return MaskData(this, event)" value="<%=dfilho3%>"></td>
                  </tr>
                  <tr>
                    <td width="773" height="30" colspan="3">
<input name="nfilho4" type="text" class="form" id="nfilho4" size="40" onKeyUp="javascript:alta(this);" value="<%=nfilho4%>">
<input name="dfilho4" type="text" class="form" id="dfilho4" onKeyUp="javascript:alta(this);" size="10" maxlength="10" onKeyPress="return MaskData(this, event)" value="<%=dfilho4%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Revendo 
						Natura, Avon e outros:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
<input name="revendooutros" type="checkbox" id="revendooutros" value="Sim" <%if revendooutros = "Sim" then%>checked<%end if%>></td>
                  </tr>
                  <tr>
                    <td width="184" height="30">&nbsp;
              </td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="184" height="30">
              <input name="Reset" type="reset" class="botaoform" id="button" value=" Limpar ">&nbsp;&nbsp;<input name="button2" type="submit" class="botaoform" id="button2" value=" Cadastrar "></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="184" height="30">&nbsp;</td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">&nbsp;</td>
                  </tr>
                </table>
                </form>
                </div></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "cadastrar.ok"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="center" valign="top">
              <br>
              <br>
              <br>
              <br>
              <br>
              <br>
              <font color="#F8D000"><b>Cadastro realizado com sucesso!</b></font></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "passarpedido"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top"><div id="text_int"><h1>
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
              <form action="pedidos.asp?" method="POST" onSubmit="return (FormOk(this) &amp;&amp; checa(this))">
<script language="javascript" src="js/functions.js"></script>
<script language="javascript" src="js/scripts.js"></script>
<script type="text/javascript">
function handleEnter (field, event) {
		var keyCode = event.keyCode ? event.keyCode : event.which ? event.which : event.charCode;
		if (keyCode == 13) {
			var i;
			for (i = 0; i < field.form.elements.length; i++)
				if (field == field.form.elements[i])
					break;
			i = (i + 1) % field.form.elements.length;
			field.form.elements[i].focus();
			return false;
		} 
		else
		return true;
	}    
function autoMove (field, event) {
		var keyCode = event.keyCode ? event.keyCode : event.which ? event.which : event.charCode;
		var i;
			for (i = 0; i <= field.form.elements.length; i++)
				if (field == field.form.elements)
					break;
			i = (i + 1) % field.form.elements.length;
			field.form.elements.focus();
			return false;
}
function Cancelar() {
	if(confirm('Tem certeza que deseja cancelar o Pedido?')) {
	    location.href = "?action=cancelarpedido";
	 	return true;
	}
}
                              </script>
<script type="text/javascript" src="js/jquery.js"></script>
<script type="text/javascript" src="js/jquery.validate.js"></script>
<script type="text/javascript" src="js/jquery_003.js"></script> 
<script type="text/javascript">
	$(function() {
		var offset = $("#floatDiv").offset();
		var topPadding = 15;
		$(window).scroll(function() {
			if ($(window).scrollTop() > offset.top) {
				$("#floatDiv").stop().animate({
					marginTop: $(window).scrollTop() - offset.top + topPadding
				});
			} else {
				$("#floatDiv").stop().animate({
					marginTop: 0
				});
			};
		});
	});
          </script>
            <style>
			#floatDiv{
				position: absolute;
				left: 605px;
				border: 1px solid #8A0608; 
				background-color: #FFF;
				margin: 0 auto;
				width: 140px;
			}
			#left-div {	position: relative; }
		                  </style>
			<div id="left-div" height="100%">
			<div id="floatDiv" align="center">
			<div style="padding: 10px;"><font size="1" color="#8A0608"><strong>
				VALOR TOTAL DO SEU PEDIDO</strong></font><br />
			</div>
			<input name="str_totalGeral" value="R$ 0,00" size="7" dir="rtl" onKeyUp="formatValueNumber(this, '##.##0,00', event)" disabled="disabled" style="background-color:#FFF; font-size: 22px; padding: 5px; border-color: #FFF; border-style: double; float:right" />
			</div></div>
              <div id="conteudo-formulario">
          <div id="tudo">
          <input type="hidden" name="action" value="enviarpedido_2">
          <div id="TabbedPanels1" class="TabbedPanels">
            <ul class="TabbedPanelsTabGroup">
              <li class="TabbedPanelsTab TabbedPanelsTabSelected" tabindex="0">
				Formulário de Pedidos</li>
              <li class="TabbedPanelsTab" tabindex="0">Informações de Pagamento</li>
            </ul>
            <div class="TabbedPanelsContentGroup">
              <div class="TabbedPanelsContent TabbedPanelsContentVisible" style="display: block; ">
<table width="560" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tbody><tr class="textopedido">
                    <td align="center">Quant.</td>
                    <td align="center">Refer.</td>
                    <td align="center">Tam.</td>
                    <td align="center">Descrição</td>
                    <td align="center">Página</td>
                    <td align="center">Cor</td>
                    <td align="center">Preço Unit.</td>
                    <td align="center">Preço Tot.</td>
                    </tr>

<%
for i = 1 to 60
%>
					<tr>
                    <td height="25" align="center"><input name="str_qtd<%=i%>" type="text" class="formpedido" id="str_qtd<%=i%>" size="2" maxlength="2" onKeyUp="formatValueNumber(this, '##', event);nextField(this, event)" onBlur="total(<%=i%>, 300)" onKeyPress="return Nquant(this, event)"></td>
                    <td height="25" align="center"><input name="str_referencia<%=i%>" type="text" class="formpedido" id="str_referencia<%=i%>" onkeypress="return handleEnter(this, event)" onkeypress="return Referencia(this, event)" size="10" maxlength="10"></td>
                    <td height="25" align="center"><input name="str_tamanho<%=i%>" type="text" class="formpedido" id="str_tamanho<%=i%>" size="3" maxlength="3" onKeyPress="return handleEnter(this, event)" onKeyUp="javascript:alta(this);"></td>
                    <td height="25" align="center"><input name="str_descricao<%=i%>" type="text" class="formpedido" id="str_descricao<%=i%>" size="28" maxlength="40" onKeyPress="return handleEnter(this, event)" onKeyUp="javascript:alta(this);"></td>
                    <td height="25" align="center"><input name="str_pagina<%=i%>" type="text" class="formpedido" id="str_pagina<%=i%>" size="3" maxlength="3" onKeyPress="return handleEnter(this, event);return Nseg(this, event)"></td>
                    <td height="25" align="center"><input name="str_cor<%=i%>" type="text" class="formpedido" id="str_cor<%=i%>" size="8" maxlength="8" onKeyUp="javascript:alta(this);" onKeyPress="return Cor(this, event)"></td>
                    <td align="center"><input name="str_unitario<%=i%>" dir="rtl" type="text" class="formpedido" id="str_unitario<%=i%>" size="8" maxlength="7" onKeyUp="formatValueNumber(this, '#.##0,00', event)" onBlur="total(<%=i%>, 300)"></td>
                    <td height="25" align="center"><input name="str_total<%=i%>" dir="rtl" type="text" class="formpedido" id="str_total<%=i%>" size="8" maxlength="7" onKeyUp="formatValueNumber(this, '##.##0,00', event)" onFocus="autoMove(this, event)"></td>
                    </tr>
<%
next
%>
                </tbody>
                </table>
                </div>
              <div class="TabbedPanelsContent" style="display: none; ">
<table width="580" border="0" align="center" cellpadding="0" cellspacing="0" height="536">
                  <tbody><tr>
                    <td height="40" colspan="4" class="textonoticias">
					INFORMAÇÕES DA CONSULTORA</td>
                    </tr>
                  <tr>
                    <td width="65" height="25"><span class="textocomun">Nome:</span></td>
                    <td height="25" colspan="3"><input name="nome" type="text" class="formpedido" id="nome" size="65" erro="Preencha o seu Nome" onKeyUp="javascript:alta(this);"></td>
                    </tr>
                  <tr>
                    <td height="25"><span class="textocomun">E-mail:</span></td>
                    <td height="25" colspan="3"><input name="email" type="text" class="formpedido" id="email" size="30" erro="Preencha o seu E-mail" onKeyUp="javascript:alta(this);"></td>
                  </tr>
                  <tr>
                    <td height="25"><span class="textocomun">Endereço:</span></td>
                    <td height="25" colspan="3"><input name="endereco" type="text" class="formpedido" id="endereco" size="40" erro="Preencha o seu Endereço" onKeyUp="javascript:alta(this);"></td>
                    </tr>
                  <tr>
                    <td height="25"><span class="textocomun">Bairro:</span></td>
                    <td width="257" height="25"><input name="bairro" type="text" class="formpedido" id="bairro" size="38" onKeyUp="javascript:alta(this);"></td>
                    <td width="57" height="25" class="textocomun">Código:</td>
                    <td width="201" height="25">
                    <input name="codigo" type="text" class="formpedido" id="codconsult" size="14" maxlength="14"></td>
                  </tr>
                  <tr>
                    <td height="25"><span class="textocomun">Cidade:</span></td>
                    <td height="25"><input name="cidade" type="text" class="formpedido" id="cidade" size="35" onKeyUp="javascript:alta(this);"></td>
                    <td height="25"><span class="textocomun">Telefone:</span></td>
                    <td height="25">
                    <input name="tel" type="text" class="formpedido" id="telefone" size="14" maxlength="14" erro="Preencha o seu Telefone" onKeyPress="return MaskTelefone(this, event)"></td>
                  </tr>
                  <tr>
                    <td height="40" colspan="4" class="textonoticias">MARQUE A 
					SUA OPÇÃO DE PAGAMENTO:           
</td>
                  </tr>
                  <tr>
                    <td colspan="4" height="66"><span class="textocomun">
                      <label>
                        <input type="radio" name="pagamento" value="credhermes">
                        Credhermes</label>
                    </span>
                      <label><span class="textonoticias"> (Crédito direto da 
					Hermes, até 14 dias para pagar após a entrega da caixa)</span></label>
                      <span class="textocomun">
                      <label><br>
                      </label>
                      <input type="radio" name="pagamento" value="cartao">
					Cartão de Crédito<br>
<label>
  <input type="radio" name="pagamento" value="cliente">
  					Cartão de Crédito</label>
					da Cliente                    </span></td>
                  </tr>
                  <tr>
                    <td height="25" colspan="4" class="textocomun">Número do 
					Cartão:
                      <input name="numerocard" type="text" class="formpedido" id="ncartao" size="20" maxlength="19" onKeyPress="return Ncartao(this, event)"></td>
                  </tr>
                  <tr>
                    <td height="25" colspan="4"><span class="textocomun">Cód. de 
					Segurança:
                        <input name="codigoseg" type="text" class="formpedido" id="nseg" size="4" maxlength="4" onKeyPress="return Nseg(this, event)">
                      <span class="textonoticias">três ou quartro últimos 
					números no verso do Cartão</span></span></td>
                  </tr>
                  <tr>
                    <td height="25" colspan="4"><span class="textocomun">
					Validade do Cartão: Mês/Ano
                      <select name="mes" class="formpedido" id="nmes">
                        <option value="Janeiro">Janeiro</option>
                        <option value="Fevereiro">Fevereiro</option>
                        <option value="Março">Março</option>
                        <option value="Abril">Abril</option>
                        <option value="Maio">Maio</option>
                        <option value="Junho">Junho</option>
                        <option value="Julho">Julho</option>
                        <option value="Agosto">Agosto</option>
                        <option value="Setembro">Setembro</option>
                        <option value="Outubro">Outubro</option>
                        <option value="Novembro">Novembro</option>
                        <option value="Dezembro">Dezembro</option>
                      </select>
<input name="ano" type="text" class="formpedido" id="nano" onKeyPress="return Nano(this, event)" size="4" maxlength="4">
                    </span><span class="textonoticias">(9999)</span></td>
                  </tr>
                  <tr>
                    <td height="25" colspan="4" class="textocomun">Número de 
					Parcelas: 
                      <select name="parcelas" class="formpedido" id="parcelas">
                        <option value="1">1</option>
                        <option value="2">2</option>
                        <option value="3">3</option>
                        <option value="4">4</option>
                        <option value="5">5</option>
                        <option value="6">6</option>
                        <option value="7">7</option>
                        <option value="8">8</option>
                        <option value="9">9</option>
                        <option value="10">10</option>
                      </select></td>
                  </tr>
                  <tr>
                    <td height="25" colspan="4" class="textocomun">Nome da 
					Cliente: 
                      <input name="nometitular" type="text" class="formpedido" id="cliente" size="35" onKeyUp="javascript:alta(this);">                    
					Telefone da Cliente: 
                      <input name="telc" type="text" class="formpedido" id="tcliente" size="14" maxlength="14" onKeyPress="return MaskTelefone(this, event)"></td>
                  </tr>
                  <tr>
                    <td height="40" colspan="4" class="textonoticias">PROGRAMA 
					PONTO VAI PONTO VEM</td>
                  </tr>
                  <tr class="textocomun">
                    <td height="25" colspan="4">Código do Produto:
                      <input name="codigopvpv" type="text" class="formpedido" id="codigopvpv" size="9" maxlength="9"></td>
                    </tr>
                  <tr>
                    <td height="25" colspan="4" class="textocomun">Nome do 
					Produto:
                      <input name="produtopvpv" type="text" class="formpedido" id="produtopvpv" size="35" onKeyUp="javascript:alta(this);"></td>
                    </tr>
                  <tr>
                    <td height="50" colspan="2" align="center">
                    <input name="button2" type="reset" class="botaoform" id="button2" value=" Limpar Pedido "></td>
                    <td colspan="2" height="50">
                    <input name="button" type="submit" class="botaoform" id="button" value=" Enviar Pedido "></td>
                    </tr>
                </tbody></table>
            </div>
            </div>
          </div>
          </div>
<script type="text/javascript">
//<![CDATA[

onMaxLengthFocusNext = function(){
function aE( o, e, h ) { o.addEventListener ? o.addEventListener( e, h, false ) : o.attachEvent ? o.attachEvent( 'on' + e, h ) : o[ 'on' + e ] = h; };
function kU( e ){
 var k = ( e = e || event ).which || e.keyCode || 0, e = e.target || e.srcElement, el = e.form.elements;
 if ( e.value.length >= ( e.getAttribute( 'maxlength' ) || e.value.length+1 ) && /[\wÀ-ÿ ]/.test( String.fromCharCode( k ) ) ){
  for ( k = el.length; el[--k] != e; );
  return el[ ( ++k ) * ( k < el.length ) ].focus();
 }
 return true;
}
for( var d, f = ( d = document ).forms, i = -1; ++i < d.forms.length; aE( f[i], 'keyup', kU ) );
}();

//]]>
var TabbedPanels1 = new Spry.Widget.TabbedPanels("TabbedPanels1");
                              </script>
            </form>
		      	</td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "pedidos.ok"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="center" valign="top">
              <br>
              <br>
              <br>
              <br>
              <br>
              <br>
              <font color="#F8D000"><b>Querida consultora, você receberá em breve, um TORPEDO através do seu celular, confirmando a chegada do seu pedido com sucesso.<br /><br />
OBS: Mantenha-nos informados do número do seu celular atualizado, para que possamos mantê-la informada.
</b></font></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "pecacatalogos"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Peça catalogos</h1>
                <p align="justify"><strong>Preencha o formulário abaixo, pra 
				você começar a receber os catálogos HERMES e BELLA na sua casa.</strong></p>
                <form action="?">
                <input type="hidden" name="action" value="catalogos" size="20">
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="773">
                  <tr>
                    <td width="773" height="30" colspan="3">
                    <div id="text_int"><i><b>Obs.:<br>
                      </b>- Todos os campos marcados com<font color="#F8D000"><b> 
						*
                      </b></font>são obrigatórios.<br>
                      	- Para efetuar o preenchimento dos campos, CEP, 
						DDD+Telefone e DDD+Celular digite somente os números, 
						sem pontuação ou espaços.</i></div><br /></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Nome 
						Completo:
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="nome" type="text" class="form" id="nome" size="42" erro="Seu Nome" onKeyUp="javascript:alta(this);" value="<%=nome%>">
              <font color="#F8D000" size="2"><b><%=session("erro_nome")%></b></font></td>
                  </tr>
                  <tr>
                    <td width="184" height="30">
                    <div id="text_int">Endereço: <b><font color="#F8D000">*</font></b></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="endereco" type="text" class="form" id="endereco" size="42" erro="Seu Endereço" onKeyUp="javascript:alta(this);" value="<%=endereco%>">
              <font color="#F8D000" size="2"><b><%=session("erro_endereco")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Bairro: <b>
                      <font color="#F8D000">*</font></b></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30"> 
              <input name="bairro" type="text" class="form" id="bairro" size="32" erro="Seu Bairro" onKeyUp="javascript:alta(this);" value="<%=bairro%>">
              <font color="#F8D000" size="2"><b><%=session("erro_bairro")%></b></font>
					</td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">CEP: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="cep" type="text" class="form" id="cep" size="10" maxlength="8" erro="Seu CEP" onKeyPress="return CEP(this, event)" value="<%=cep%>">
              <font color="#F8D000" size="2"><b><%=session("erro_cep")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Cidade: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="cidade" type="text" class="form" id="cidade" size="32" erro="Sua Cidade" onKeyUp="javascript:alta(this);" value="<%=cidade%>">
              <font color="#F8D000" size="2"><b><%=session("erro_cidade")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Estado:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
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
</option><option value="SERGIPE" <%if estado = "SERGIPE" then%>selected="selected"<%end if%>>
SERGIPE
</option><option value="TOCANTINS" <%if estado = "TOCANTINS" then%>selected="selected"<%end if%>>
TOCANTINS
		  </option></select></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">DDD+Telefone: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="tel" type="text" class="form" id="tel" size="14" maxlength="10" value="<%=tel%>">
              <font color="#F8D000" size="2"><b><%=session("erro_tel")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">DDD+Celular:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
<input name="cel" type="text" class="form" id="cel" size="14" erro="Seu Telefone" maxlength="10" value="<%=cel%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int" style="width: 184; height: 16">
                      E-mail:<font color="#F8D000"><b> *</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="email" type="text" class="form" id="email" size="35" value="<%=email%>">
              <font color="#F8D000" size="2"><b><%=session("erro_email")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30">&nbsp;
              </td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="184" height="30">
              <input name="Reset" type="reset" class="botaoform" id="button" value=" Limpar ">&nbsp;&nbsp;<input name="button2" type="submit" class="botaoform" id="button2" value=" Enviar "></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="184" height="30">&nbsp;</td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">&nbsp;</td>
                  </tr>
                </table>
                </form>
                </div></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "contato"
        %>
         <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Contato</h1>
                <p align="justify"><strong>Preencha o formulário abaixo, para 
				falar diretamente com a Catálogos Hermes.</strong></p>
                <form action="?">
                <input type="hidden" name="action" value="contact" size="20">
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="773">
                  <tr>
                    <td width="773" height="30" colspan="3">
                    <div id="text_int"><i><b>Obs.:<br>
                      </b>- Todos os campos marcados com<font color="#F8D000"><b> 
						*
                      </b></font>são obrigatórios.<br>
                      	- Para efetuar o preenchimento dos campos, CEP, 
						DDD+Telefone e DDD+Celular digite somente os números, 
						sem pontuação ou espaços.</i></div><br /></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Nome 
						Completo:
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="nome" type="text" class="form" id="nome" size="42" erro="Seu Nome" onKeyUp="javascript:alta(this);" value="<%=nome%>">
              <font color="#F8D000" size="2"><b><%=session("erro_nomec")%></b></font></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Cidade: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="cidade" type="text" class="form" id="cidade" size="32" erro="Sua Cidade" onKeyUp="javascript:alta(this);" value="<%=cidade%>">
              <font color="#F8D000" size="2"><b><%=session("erro_cidadec")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Estado: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
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
                    <td width="184" height="30"><div id="text_int">DDD+Telefone: 
                      <font color="#F8D000"><b>*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="tel" type="text" class="form" id="tel" size="14" maxlength="10" value="<%=tel%>">
              <font color="#F8D000" size="2"><b><%=session("erro_telc")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">DDD+Celular:</div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
<input name="cel" type="text" class="form" id="cel" size="14" erro="Seu Telefone" maxlength="10" value="<%=cel%>"></td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Email:<font color="#F8D000"><b> 
						*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
              <input name="email" type="text" class="form" id="email" size="35" value="<%=email%>">
              <font color="#F8D000" size="2"><b><%=session("erro_emailc")%></b></font>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30"><div id="text_int">Assunto:<font color="#F8D000"><b> 
						*</b></font></div></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">
					<select name="assunto" class="form" id="assunto">
							<option value="Quero ser consultor(a)" <%if assunto = "Quero ser consultor(a)" then%>selected="selected"<%end if%>>
							Quero ser consultor(a)</option>
							<option value="Trocas" <%if assunto = "Trocas" then%>selected="selected"<%end if%>>
							Trocas</option>
							<option value="Dúvidas gerais" <%if assunto = "Dúvidas gerais" then%>selected="selected"<%end if%>>
							Dúvidas gerais</option>
							<option value="Dúvidas sobre pedidos" <%if assunto = "Dúvidas sobre pedidos" then%>selected="selected"<%end if%>>
							Dúvidas sobre pedidos</option>
							<option value="Outros assuntos" <%if assunto = "Outros assuntos" then%>selected="selected"<%end if%>>
							Outros assuntos</option>
					</select>
					</td>
                  </tr>
                  <tr>
                    <td width="184" height="90" valign="top">
                    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                      <tr>
                        <td width="100%" height="30"><div id="text_int">
                          Mensagem:<font color="#F8D000"><b> *</b></font></div></td>
                      </tr>
                    </table>
                    </td>
                    <td width="9" height="90">&nbsp;</td>
                    <td width="580" height="90">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="53%"><textarea rows="5" name="mensagem" cols="35"><%=mensagem%></textarea></td>
                  <td width="47%" valign="top"><font color="#F8D000" size="2"><b><%=session("erro_mensagemc")%></b></font></td>
                </tr>
              </table>
              </td>
                  </tr>
                  <tr>
                    <td width="184" height="30">&nbsp;
              </td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="184" height="30">
              <input name="Reset" type="reset" class="botaoform" id="button" value=" Limpar ">&nbsp;&nbsp;<input name="button2" type="submit" class="botaoform" id="button2" value=" Enviar "></td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="184" height="30">&nbsp;</td>
                    <td width="9" height="30">&nbsp;</td>
                    <td width="580" height="30">&nbsp;</td>
                  </tr>
                </table>
                </form>
                </div></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "contato.ok"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="center" valign="top">
              <br>
              <br>
              <br>
              <br>
              <br>
              <br>
              <font color="#F8D000"><b>Mensagem enviado com sucesso!</b></font></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "catalogos.ok"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="center" valign="top">
              <br>
              <br>
              <br>
              <br>
              <br>
              <br>
              <font color="#F8D000"><b>Pedido enviado com sucesso!</b></font></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "clubeamizade"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="4"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="613" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Clube da Amizade</h1>
                <p align="justify">
				<img border="0" src="image/ca.jpg" width="613" height="972"></div>
			  </td>
              <td width="160" align="right" valign="top">
              
              <img border="0" src="image/banner001.jpg"></td>
            </tr>
            <tr>
              <td width="100%" colspan="4">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "premiofantastico"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="4"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="613" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Prêmio Fantástico</h1>
                <p align="justify">
				<img border="0" src="image/pf.jpg" width="615" height="1149"></div>
			  </td>
              <td width="160" align="right" valign="top">
              
              <img border="0" src="image/banner001.jpg"></td>
            </tr>
            <tr>
              <td width="100%" colspan="4">&nbsp;</td>
            </tr>
          </table>
        </div>

        <%
        case "clubedascampeas"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="4"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="613" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Clube das Campeãs</h1>
                <p align="justify">
				<img border="0" src="image/cc1.jpg" width="615" height="482"><br>
                <img border="0" src="image/cc2.jpg" width="615" height="406"><br>
                <img border="0" src="image/cc3.jpg" width="615" height="380"><br>
&nbsp;</div>
			  </td>
              <td width="160" align="right" valign="top">
              
              <img border="0" src="image/banner001.jpg"></td>
            </tr>
            <tr>
              <td width="100%" colspan="4">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "datacampanha"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="4"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
               <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="613" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Data das campanhas</h1>
                <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                  <tr>
                    <td width="100%" colspan="2" align="center">
                    <div id="text_int">
                    <b>
                    <a href="http://www.catalogoshermes.com.br">
                    www.catalogoshermes.com.br</a></b><br>
                    <b>Somente para pedidos:</b> 0800<br>
                    <b>Para informações:</b> (22) 27364677<br>
                    <br></div></td>
                  </tr>
                  <%
                  Set datas = objConn.Execute("SELECT top 3 * FROM datas ORDER BY id")
                  while not datas.eof
                  %>
                  <tr>
                    <td width="100%" align="center" colspan="2" bgcolor="#8A0608"><div id="text_int"><h1>
                      <span style="font-size: 10pt">&nbsp;</span><br>
                      <%=datas("mes")%></h1></div></td>
                  </tr>
                  <tr>
                    <td width="50%" align="center" valign="top"><div id="text_int">
                      <u><b><br>
                      Passar pedido</b></u><br>
                    <br>
                      De <%=datas("1dataida")%><br />
                      De <%=datas("2dataida")%><br />
                      De <%=datas("3dataida")%><br />
                      De <%=datas("4dataida")%><br />                    
                      </div></td>
                    <td width="50%" align="center" valign="top"><div id="text_int">
                      <u><b><br>
                      Chegada</b></u><br>
                    <br>
                      Até <%=datas("1datavolta")%><br />
                      Até <%=datas("2datavolta")%><br />
                      Até <%=datas("3datavolta")%><br />
                      Até <%=datas("4datavolta")%><br />
                      <br /></div></td>
                  </tr>
                  <tr>
                    <td width="50%" height="30">&nbsp;</td>
                    <td width="50%" height="30">&nbsp;</td>
                  </tr>
                  <%
                  datas.movenext
                  Wend
				  datas.close
				  set datas = nothing
                  %>
                </table>
                </div>
              </td>
              <td width="160" align="right" valign="top">
              <img border="0" src="image/banner001.jpg"></td>
            </tr>
            <tr>
              <td width="100%" colspan="4">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "pontovaipontovem"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="773">
                <tr>
                  <td width="600" valign="top"><div id="text_int"><h1>Programa 
					Ponto Vai, Ponto Vem</h1>
                    <p align="justify"><img border="0" src="image/pvpv01.jpg"><br>
                    <img border="0" src="image/pvpv02.jpg"><br>
                    <img border="0" src="image/pvpv03.jpg"></div></td>
                  <td width="173" height="461" align="right" valign="top">
                  <img border="0" src="image/banner001.jpg"></td>
                </tr>
                <tr>
                  <td width="100%" colspan="2">&nbsp;</td>
                </tr>
                <tr>
                  <td width="100%" colspan="2">&nbsp;</td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "superindiqueeganhe"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="773">
                <tr>
                  <td width="600" valign="top"><div id="text_int"><h1>Super 
					Indique e Ganhe</h1>
                    <p align="justify"><img border="0" src="image/sp1.jpg"><br>
                    <img border="0" src="image/sp2.jpg"><br>
                    <img border="0" src="image/sp3.jpg"></div></td>
                  <td width="173" height="461" align="right" valign="top">
                  <img border="0" src="image/banner001.jpg"></td>
                </tr>
                <tr>
                  <td width="100%" colspan="2">&nbsp;</td>
                </tr>
                <tr>
                  <td width="100%" colspan="2">&nbsp;</td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "superindiqueeganhemaster"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="773">
                <tr>
                  <td width="600" valign="top"><div id="text_int"><h1>Super 
					Indique e Ganhe para Master</h1>
                    <p align="justify"><img border="0" src="image/sigm1.jpg"><br>
                    <img border="0" src="image/sigm2.jpg"><br>
                    </div></td>
                  <td width="173" height="461" align="right" valign="top">
                  <img border="0" src="image/banner001.jpg"></td>
                </tr>
                <tr>
                  <td width="100%" colspan="2">&nbsp;</td>
                </tr>
                <tr>
                  <td width="100%" colspan="2">&nbsp;</td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "fidelidadepremiada"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="773">
                <tr>
                  <td width="600" valign="top"><div id="text_int"><h1>Fidelidade 
					Premiada</h1>
                    <p align="justify"><br>
                    <img border="0" src="image/fp1.jpg" width="588" height="453"><br>
                    <img border="0" src="image/fp2.jpg" width="588" height="450"><br>
                    <img border="0" src="image/fp3.jpg" width="588" height="463"><br>
                    &nbsp;</div></td>
                  <td width="173" height="461" align="right" valign="top">
                  <img border="0" src="image/banner001.jpg"></td>
                </tr>
                <tr>
                  <td width="100%" colspan="2">&nbsp;</td>
                </tr>
                <tr>
                  <td width="100%" colspan="2">&nbsp;</td>
                </tr>
              </table>
              </td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>

        <%
        case "comofunciona"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="4"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              </td>
              <td width="15">&nbsp;
              </td>
              <td width="613" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Como funcionamos???</h1>
                <p align="justify"><b>O seu pedido poderá ser enviado de três 
				formas:<br>
                </b><br>
                <b>1. </b>Poderá enviar os seus pedidos através de nosso 0800.<br>
                Sua ligação será gratuita e exclusiva para enviar os seus 
				pedidos. <br>
                Para informações ligue para (22) 27364677.<br>
                <b>2. </b>Poderá passar diretamente na loja.<br>
                <b>3. </b>Ou poderá passar pelo nosso site:
                <a href="http://www.catalogoshermes.com.br">
                www.catalogoshermes.com.br</a><br>
                <br>
                <br>
                <b>A sua caixa:<br>
                </b><br>
                Será entregue na sua casa (em qualquer lugar).<br>
                <br>
                <br>
                <b>Pedido mínimo:<br>
                </b><br>
                O pedido mínimo é de R$ 99,00 somando os catálogos 
				individualmente.<br>
                Fazemos troca de tamanho<br>
                <br>
                <br>
                <b>Prazo de pagamento:<br>
                </b><br>
                A consultora Hermes terá um prazo de 12 a 14 dias para pegar o 
				seu boleto, após a entrega da sua caixa. Ou ainda, a consultora 
				Hermes ou a sua clientes, poderão pagar parceladamente em até 
				10X sem juros, no seu cartão de crédito.<br>
                <br>
                <br>
                <b>E muito mais:<br>
                </b><br>
                Você poderá ganhar brindes em todos os seus pedidos e ainda 
				acumular pontos em todos os seus pedidos para trocar por PRÊMIOS 
				no programa &quot;Ponto vai ponto vem&quot; da Hermes.<br>
                <br>
&nbsp;</div>
			  </td>
              <td width="160" align="right" valign="top">
			      <img border="0" src="image/banner001.jpg"></td>
            </tr>
            <tr>
              <td width="100%" colspan="4">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "cartas"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="4"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              </td>
              <td width="15">&nbsp;
              </td>
              <td width="613" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Critérios de envio de cartas de aniversário 
				Consultora, Mãe, Filhos e Marido</h1>
                <b>Para GANHAR o prêmio:<br>
                </b><br>
                - A consultora precisa ter realizado pedidos nos últimos 12 
				meses;<br>
                - Não pode estar inadimplente e com nome no SPC;<br>
                - Não pode estar bloqueada por excesso de devolução;<br>
                - Não pode ter APENAS 1 compra e 1 devolução no seu histórico;<br>
                - Para receber cartas de<b> PAI</b>, <b>MÃE</b>, <b>FILHOS</b> e 
                <b>MARIDO</b>, precisa estar cadastrada no Clube da Amizade 
				ANTES do envio da carta;<br>
                - Carta de aniversário de filhos só é enviada para quem tem 
				filhos de até 14 anos;<br>
                - Último canal de compras: <b>DISTRIBUIDOR ATIVO</b>;<br>
                - Ter data de aniversáro no cadastro;<br>
                - Endereços com status de &quot;Carta devolvida pelos Correios&quot; serão 
				excluídos da listagem de envio de cartas de aniversário;<br>
                - Fazer pedido com o código de cliente<b> INFORMADO NA CARTA</b>;<br>
                - Respeitar a data de validade informada na carta;<br>
                - Fazer o pedido mínimo de:<br>
&nbsp; R$ 99,00 - Cartas de Aniversário;<br>
&nbsp; R$ 75,00 - Cartas para inativas.
                </div>
              </td>
              <td width="160" align="right" valign="top">
			      <img border="0" src="image/banner001.jpg"></td>
            </tr>
            <tr>
              <td width="100%" colspan="4">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "ganhemais"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="4"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              </td>
              <td width="15">&nbsp;
              </td>
              <td width="613" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Ganhe Mais Pontos - Objetivos</h1>
                Prezado(a),<br>
                <br>
                Todos os meses, informamos os objetivos de vendas para as 
				Consultoras, no formulário de pedidos que segue na caixa de 
				mercadorias.<br>
                Estes objetivos sempre são para o total de vendas do mês 
				seguinte.<br>
                Esta é a promoção &quot;Ganhe Mais Pontos&quot;<br>
                <br>
                Com esta promoção conquistamos mais pedidos, as Consultoras 
				ganham pontos extras na Ponto Vai Ponto Vem e ainda aumentam os 
				seus lucros e os seus pontos!<br>
                <br>
                <b>Pontos extras:</b><br>
                <br>
                - Consultora <b>Ouro</b> Ganha Mais<b> 500 pontos extras </b>na 
				Ponto vai Ponto Vem.<br>
                <br>
                - Consultora <b>Prata </b>Ganha Mais <b>200 pontos extras </b>na 
				Ponto vai Ponto Vem.<br>
                <br>
                - Consultora<b> Especial</b> Ganha Mais<b> 100 pontos extras </b>
                na Ponto vai Ponto Vem.<br>
                <br>
                - Consultora<b> Não Cadastrada</b> Ganha Mais<b> 500 pontos 
				extras </b>na Ponto vai Ponto Vem.<br>
                <br>
                Vejam os seus objetivos na Ganhe Mais Pontos para este mês, ou 
				consulte o seu Distribuidor Hermes.<br>
                Ligue ou visite suas clientes, para que você possa aproveitar de 
				todas as vantagens que a Hermes te proporciona.<br>
&nbsp;</div>
              </td>
              <td width="160" align="right" valign="top">
                  <img border="0" src="image/banner001.jpg"></td>
            </tr>
            <tr>
              <td width="100%" colspan="4">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        case "noticias"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
        </div>
        <%
        case "lucros"
        %>
        <div id="text_int">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <tr>
              <td width="100%" height="20" colspan="3"><span style="font-size: 1pt">
				&nbsp;</span></td>
            </tr>
            <tr>
              <td width="200" align="left" valign="top">
              <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="100%" height="24" background="image/topmenu.png" colspan="2">&nbsp;
					</td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=comofunciona">
					Como Funcionamos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=cadastre-se">
					Fazer Cadastro</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=passarpedido">
					Passar Pedidos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pecacatalogos">
					Peça Catálogos Aqui</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=datacampanha">
					Data das Campanhas</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubedascampeas">
					Clube das Campeãs</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=clubeamizade">
					Clube da Amizade</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu"><a href="?hl=pontovaipontovem">
					Ponto vai Ponto Vem</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="25" align="center">
                  <img border="0" src="image/marcador.png"></td>
                  <td width="170" bgcolor="#8A0608" height="25"><div id="menu">
					<a href="?hl=fidelidadepremiada">Fidelidade Premiada</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="#">
					Seja Uma Promotora Master</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <span style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</span></td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=ganhemais">
					Ganhe Mais Pontos Objetivos</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="45" align="center">
                  <img border="0" src="image/marcador.png"><br>
&nbsp;</td>
                  <td width="170" bgcolor="#8A0608" height="45"><div id="menu"><a href="?hl=cartas">
					Carta de Aniversário Critérios de Envio</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="30" bgcolor="#8A0608" height="60" align="center">
                  <font style="font-size: 9pt">
                  <img border="0" src="image/marcador.png"><br>
                  <br>
&nbsp;</font></td>
                  <td width="170" bgcolor="#8A0608" height="60"><div id="menu"><a href="?hl=lucros">
					Lucros dos Catálogos Ofertas e Promoções do Mês</a></div></td>
                </tr>
                <tr>
                  <td width="200" bgcolor="#8A0608" height="3" align="center" colspan="2" background="image/pontilhado.png">
                  <span style="font-size: 1pt">&nbsp;</span></td>
                </tr>
                <tr>
                  <td width="100%" background="image/rodmenu.png" height="24" colspan="2">&nbsp;
					</td>
                </tr>
              </table>              </td>
              </td>
              <td width="15">&nbsp;
              </td>
              <td width="773" align="left" valign="top">
              <div id="text_int">
                <h1 align="justify">Lucros dos catálogos ofertas e promoções do 
				mês</h1>
                </div></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
          </table>
        </div>
        <%
        end select
        %>
        </td>
      </tr>
    </table>
    <p><span style="font-size: 1pt">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="100%" height="7" align="center" bgcolor="#8A0608"><span style="font-size: 1pt">
	&nbsp;</span></td>
  </tr>
  <tr>
    <td width="100%" height="115" align="center" valign="top">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="988" height="115">
      <tr>
        <td align="left" valign="bottom" width="182" height="15">
        <span style="font-size: 1pt">&nbsp;</span></td>
        <td align="left" valign="bottom" width="630" height="15">
        </td>
        <td align="right" width="176" valign="bottom" height="15">
        <span style="font-size: 1pt">&nbsp;</span></td>
      </tr>
      <tr>
        <td align="left" valign="top" width="182">
          <br>
			&nbsp;</div></td>
        <td align="left" valign="top" width="630"><div id="text_ext">
          <p align="center">
          <font color="#8A0608"><b>Nossa Loja</b><br>
          Fica atrás da Casa &amp; Vídeo, na galeria ao lado - Primeira loja.<br>
          Rua: Gil de Góis, 103 loja n.o 34 - Centro - Campos dos Goytacazes - 
			Rio de Janeiro.</font><font color="#A02C27"><br>
          <br>
          </font><b><font color="#8A0608">Copyright&nbsp;<%=year(Now)%> © Catálogo 
			Hermes. Todos os direitos reservados.</font></b></div></td>
        <td align="right" width="176">
        <br>
		<font size="2"><br>
		<br>
		</font><b><font color="#8A0608" size="2"><%=Application("chacessos")%></font></b></td>
      </tr>
      </table>
    </td>
  </tr>
</table>
<%
Application("fwsseg") = "fusion"
objConn.close
set objConn = nothing
%>
</body>
</html>