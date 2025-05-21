' ~~ Definindo Variáveis públicas. ~~ '
Public DataPrevistaArr As Variant
Public ProbabilidadeArr As Variant
Public StatusArr As Variant
Public CaronaArr As Variant
Public cnpjVariant As String
Public emailVariant As String
Public telefoneVariant As String
Public nomeVariant As String
Public sobrenomeVariant As String
Public orgaoVariant As String
Public dataprevistaVariant As String
Public skuVariant As String
Public customizadoVariant As String
Public qtdtotalVariant As String
Public qtdfaturadoVariant As String
Public fattotalVariant As String
Public faturadoVariant As String
Public licitacaoVariant As String
Public caronaVariant As String
Public comentariosVariant As String
Public probabilidadeVariant As String
Public statusVariant As String
Public numeropedidonfVariant As String
Public qtdtotalVariantInt As Integer
Public qtdfaturadoVariantInt As Integer
Public fattotalVariantCur As Currency
Public faturadoVariantCur As Currency
Public linhaConsultada As Integer
Public senha As String
Public CNPJCON As String
Public EMAILCON As String
Public TELEFONECON As String
Public NOMECON As String
Public SOBRENOMECON As String
Public ORGAOCON As String
Public DATAPREVISTACON As String
Public SKUCON As String
Public CUSTOMIZADOCON As String
Public QTDTOTALCON As String
Public QTDFATURADOCON As String
Public FATTOTALCON As String
Public FATURADOCON As String
Public LICITACAOCON As String
Public CARONACON As String
Public COMENTARIOSCON As String
Public PROBABILIDADECON As String
Public STATUSCON As String
Public NUMEROPEDIDONFCON As String

' ~~ Função para verificar se valor do campo está na lista suspensa. ~~ '
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean

IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)

End Function

' ~~ Função para identificar última linha preenchida em uma coluna. ~~ '
Function lastRow(worksheetName As String, coluna As String)

lastRow = ThisWorkbook.Worksheets(worksheetName).Cells(ThisWorkbook.Worksheets(worksheetName).Rows.Count, coluna).End(xlUp).Row

End Function

' ~~ Quando UserForm é aberto, importa valores das listas suspensas e arrays. ~~ '
Public Sub UserForm_Initialize()

CaronaArr = Array("SIM", "NÃO")
DataPrevistaArr = Array("Q1", "Q2", "Q3", "Q4")
ProbabilidadeArr = Array("Alta", "Média", "Baixa")
StatusArr = Array("Em Análise", "Em Homologação", "Certame Ganho", "Aguardando Empenho", "Enviado a Produção", "Finalizado/Perdido", "Finalizado/Ganho")

With Me.DATAPREVISTA
    .AddItem "Q1"
    .AddItem "Q2"
    .AddItem "Q3"
    .AddItem "Q4"
End With

With Me.PROBABILIDADE
    .AddItem "Alta"
    .AddItem "Média"
    .AddItem "Baixa"
End With

With Me.STATUS
    .AddItem "Em Análise"
    .AddItem "Em Homologação"
    .AddItem "Certame Ganho"
    .AddItem "Aguardando Empenho"
    .AddItem "Enviado a Produção"
    .AddItem "Finalizado/Perdido"
    .AddItem "Finalizado/Ganho"
End With

With Me.CARONA
    .AddItem "SIM"
    .AddItem "NÃO"
End With

End Sub

' ~~ Quando inserido novo registro. ~~ '
Public Sub INSERIR_Click()

If linhaConsultada Then
    MsgBox ("Você está re-inserindo um registro que acabou de ser consultado. Salve as alterações ao invés de inserir um novo.")
    Exit Sub
End If

senha = "senha"
ActiveSheet.Unprotect senha

' ~~ Coleta valores dos campos. ~~ '
cnpjVariant = CNPJ.Value
emailVariant = EMAIL.Value
telefoneVariant = TELEFONE.Value
nomeVariant = NOME.Value
sobrenomeVariant = SOBRENOME.Value
orgaoVariant = ORGAO.Value
dataprevistaVariant = DATAPREVISTA.Value
skuVariant = SKU.Value
customizadoVariant = CUSTOMIZADO.Value
qtdtotalVariant = QTDTOTAL.Value
qtdfaturadoVariant = QTDFATURADO.Value
fattotalVariant = FATTOTAL.Value
faturadoVariant = FATURADO.Value
licitacaoVariant = LICITACAO.Value
caronaVariant = CARONA.Value
comentariosVariant = COMENTARIOS.Value
probabilidadeVariant = PROBABILIDADE.Value
statusVariant = STATUS.Value
numeropedidonfVariant = NUMEROPEDIDONF.Value

' ~~ CNPJ. ~~ '
If cnpjVariant = "" Then
    MsgBox ("Preencha o campo: (CNPJ).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(cnpjVariant) Then
    MsgBox ("Insira apenas números no campo: (CNPJ).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Len(cnpjVariant) <> 14 Then
    MsgBox ("Insira todos os 14 dígitos no CNPJ.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ E-mail. ~~ '
If emailVariant = "" Then
    MsgBox ("Preencha o campo: (E-mail).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Telefone. ~~ '
If telefoneVariant = "" Then
    MsgBox ("Preencha o campo: (Telefone).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(telefoneVariant) Then
    MsgBox ("Insira apenas números no campo: (Telefone).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Len(telefoneVariant) < 9 Then
    MsgBox ("Insira todos os dígitos no telefone.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Nome. ~~ '
If nomeVariant = "" Then
    MsgBox ("Preencha o campo: (Nome).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Sobrenome. ~~ '
If sobrenomeVariant = "" Then
    MsgBox ("Preencha o campo: (Sobrenome).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Órgão. ~~ '
If orgaoVariant = "" Then
    MsgBox ("Preencha o campo: (Órgão).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Licitação. ~~ '
If licitacaoVariant = "" Then
    MsgBox ("Preencha o campo: (Nº Licitação).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Carona. ~~ '
If caronaVariant = "" Then
    MsgBox ("Preencha o campo: (Carona Governo).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsInArray(caronaVariant, CaronaArr) Then
    MsgBox ("Insira um valor válido no campo: (Carona Governo).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Data Prevista. ~~ '
If dataprevistaVariant = "" Then
    MsgBox ("Preencha o campo: (Data Prevista).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsInArray(dataprevistaVariant, DataPrevistaArr) Then
    MsgBox ("Insira um valor válido no campo: (Data Prevista).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ SKU. ~~ '
If skuVariant = "" Then
    MsgBox ("Preencha o campo: (SKU).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(skuVariant) Then
    MsgBox ("Insira apenas números no campo: (SKU).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Quantidade total. ~~ '
If qtdtotalVariant = "" Then
    MsgBox ("Preencha o campo: (Qtd. Peças (TOTAL)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(qtdtotalVariant) Then
    MsgBox ("Insira apenas números no campo: (Qtd. Peças (TOTAL)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
qtdtotalVariantInt = CInt(qtdtotalVariant)

' ~~ Quantidade faturado. ~~ '
If qtdfaturadoVariant = "" Then
    MsgBox ("Preencha o campo: (Qtd. Peças (FATURADO)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(qtdfaturadoVariant) Then
    MsgBox ("Insira apenas números no campo: (Qtd. Peças (FATURADO)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
qtdfaturadoVariantInt = CInt(qtdfaturadoVariant)

' ~~ Checando se quantidade faturado é maior que quantidade total. ~~ '
If qtdfaturadoVariantInt > qtdtotalVariantInt Then
    MsgBox ("Quantidade faturada não pode ser maior que quantidade total.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Faturamento total. ~~ '
If fattotalVariant = "" Then
    MsgBox ("Preencha o campo: (Valor (TOTAL)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(fattotalVariant) Then
    MsgBox ("Insira apenas números no campo: (Valor (TOTAL)). Pode ser utilizado ponto e vírgula.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
fattotalVariantCur = CCur(fattotalVariant)

' ~~ Faturado. ~~ '
If faturadoVariant = "" Then
    MsgBox ("Preencha o campo: (Valor (FATURADO)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(faturadoVariant) Then
    MsgBox ("Insira apenas números no campo: (Valor (FATURADO)). Pode ser utilizado ponto e vírgula.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
faturadoVariantCur = CCur(faturadoVariant)

' ~~ Checando se faturado total é maior que faturamento total. ~~ '
If faturadoVariantCur > fattotalVariantCur Then
    MsgBox ("Valor faturado não pode ser maior que faturamento total.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Probabilidade. ~~ '
If probabilidadeVariant = "" Then
    MsgBox ("Preencha o campo: (Probabilidade).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsInArray(probabilidadeVariant, ProbabilidadeArr) Then
    MsgBox ("Insira um valor válido no campo: (Probabilidade).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Status. ~~ '
If statusVariant = "" Then
    MsgBox ("Preencha o campo: (Status).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsInArray(statusVariant, StatusArr) Then
    MsgBox ("Insira um valor válido no campo: (Status).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ UCase. ~~ '
emailVariant = UCase(emailVariant)
nomeVariant = UCase(nomeVariant)
sobrenomeVariant = UCase(sobrenomeVariant)
orgaoVariant = UCase(orgaoVariant)
customizadoVariant = UCase(customizadoVariant)
licitacaoVariant = UCase(licitacaoVariant)
caronaVariant = UCase(caronaVariant)
comentariosVariant = UCase(comentariosVariant)

' ~~ Insere valores abaixo da última linha preenchida no PIPE. ~~ '
ultimaLinha = lastRow("PIPE", "C") + 1
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "C") = cnpjVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "D") = emailVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "E") = telefoneVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "F") = nomeVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "G") = sobrenomeVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "H") = orgaoVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "I") = dataprevistaVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "J") = skuVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "N") = customizadoVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "O") = qtdtotalVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "P") = qtdfaturadoVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "R") = fattotalVariantCur
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "S") = faturadoVariantCur
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "U") = licitacaoVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "V") = caronaVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "W") = comentariosVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "X") = probabilidadeVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "Y") = statusVariant
ThisWorkbook.Worksheets("PIPE").Cells(ultimaLinha, "Z") = numeropedidonfVariant

' ~~ Insere valores abaixo da última linha preenchida no CONSOLIDADO. ~~ '
ultimaLinha = lastRow("CONSOLIDADO", "C") + 1
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "C") = cnpjVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "D") = emailVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "E") = telefoneVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "F") = nomeVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "G") = sobrenomeVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "H") = orgaoVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "I") = dataprevistaVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "J") = skuVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "N") = customizadoVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "O") = qtdtotalVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "P") = qtdfaturadoVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "R") = fattotalVariantCur
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "S") = faturadoVariantCur
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "U") = licitacaoVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "V") = caronaVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "W") = comentariosVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "X") = probabilidadeVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "Y") = statusVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "Z") = numeropedidonfVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Range(ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "C"), ThisWorkbook.Worksheets("CONSOLIDADO").Cells(ultimaLinha, "Z")).Interior.Color = RGB(0, 255, 0)

' ~~ Ajustando colunas. ~~ '
ThisWorkbook.Worksheets("CONSOLIDADO").Columns.AutoFit
ThisWorkbook.Worksheets("PIPE").Range(ThisWorkbook.Worksheets("PIPE").Cells(3, "C"), ThisWorkbook.Worksheets("PIPE").Cells(5000, "Z")).Columns.AutoFit

' ~~ Verifica se status é Finalizado/Ganho. ~~ '
If statusVariant = "Finalizado/Ganho" Then
    If numeropedidonfVariant = "" Then
        MsgBox ("Finalizado/Ganho." & vbCrLf & vbCrLf & "Assim que tiver o pedido ou NF, não esqueça de preencher o campo.")
    End If
End If

ActiveSheet.Protect Password:=senha, AllowFiltering:=True
Unload Me

End Sub

' ~~ Consultar registro. ~~ '
Public Sub CONSULTAR_Click()

If NUMEROLINHA.Value = "" Then
    MsgBox ("Preencha o campo com a linha do registro que deseja consultar para modificar.")
    Exit Sub
Else
    linhaConsultada = NUMEROLINHA.Value
End If
If ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "C").Value = "" Or linhaConsultada = 3 Then
    MsgBox ("Registro consultado não existe.")
    Exit Sub
End If

' ~~ Insere valores nos campos do UserForm ~~ '
CNPJ.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "C").Value
EMAIL.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "D").Value
TELEFONE.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "E").Value
NOME.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "F").Value
SOBRENOME.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "G").Value
ORGAO.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "H").Value
DATAPREVISTA.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "I").Value
SKU.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "J").Value
CUSTOMIZADO.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "N").Value
QTDTOTAL.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "O").Value
QTDFATURADO.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "P").Value
FATTOTAL.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "R").Value
FATURADO.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "S").Value
LICITACAO.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "U").Value
CARONA.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "V").Value
COMENTARIOS.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "W").Value
PROBABILIDADE.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "X").Value
STATUS.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "Y").Value
NUMEROPEDIDONF.Value = ThisWorkbook.Worksheets("PIPE").Cells(linhaConsultada, "Z").Value

' ~~ Atribui valores dos campos às variáveis. ~~ '
CNPJCON = CNPJ.Value
EMAILCON = EMAIL.Value
TELEFONECON = TELEFONE.Value
NOMECON = NOME.Value
SOBRENOMECON = SOBRENOME.Value
ORGAOCON = ORGAO.Value
DATAPREVISTACON = DATAPREVISTA.Value
SKUCON = SKU.Value
CUSTOMIZADOCON = CUSTOMIZADO.Value
QTDTOTALCON = QTDTOTAL.Value
QTDFATURADOCON = QTDFATURADO.Value
FATTOTALCON = FATTOTAL.Value
FATURADOCON = FATURADO.Value
LICITACAOCON = LICITACAO.Value
CARONACON = CARONA.Value
COMENTARIOSCON = COMENTARIOS.Value
PROBABILIDADECON = PROBABILIDADE.Value
STATUSCON = STATUS.Value
NUMEROPEDIDONFCON = NUMEROPEDIDONF.Value

End Sub

' ~~ Salvar registro. ~~ '
Private Sub SALVAR_Click()

senha = "BACK72776"
ActiveSheet.Unprotect senha

If NUMEROLINHA.Value = "" Then ' <= Checa se campo de linha está em branco. '
    MsgBox ("Indique a linha do registro que foi consultado, para salvar as alterações.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
Else
    linhaModificar = NUMEROLINHA.Value
End If
If ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "C").Value = "" Then ' <= Checa que está tentando salvar existe. '
    MsgBox ("Registro está tentando ser salvo em outra linha que não existe. Favor, verificar.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If linhaConsultada <> linhaModificar Then ' <= Checa se linha consultada é a mesma modificada. '
    MsgBox ("Registro que está tentando salvar não é o mesmo que foi consultado. Favor, verificar.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Coletando dados. ~~ '
cnpjVariant = CNPJ.Value
emailVariant = EMAIL.Value
telefoneVariant = TELEFONE.Value
nomeVariant = NOME.Value
sobrenomeVariant = SOBRENOME.Value
orgaoVariant = ORGAO.Value
dataprevistaVariant = DATAPREVISTA.Value
skuVariant = SKU.Value
customizadoVariant = CUSTOMIZADO.Value
qtdtotalVariant = QTDTOTAL.Value
qtdfaturadoVariant = QTDFATURADO.Value
fattotalVariant = FATTOTAL.Value
faturadoVariant = FATURADO.Value
licitacaoVariant = LICITACAO.Value
caronaVariant = CARONA.Value
comentariosVariant = COMENTARIOS.Value
probabilidadeVariant = PROBABILIDADE.Value
statusVariant = STATUS.Value
numeropedidonfVariant = NUMEROPEDIDONF.Value

' ~~ CNPJ. ~~ '
If cnpjVariant = "" Then
    MsgBox ("Preencha o campo: (CNPJ).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(cnpjVariant) Then
    MsgBox ("Insira apenas números no campo: (CNPJ).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Len(cnpjVariant) <> 14 Then
    MsgBox ("Insira todos os 14 dígitos no CNPJ.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ E-mail. ~~ '
If emailVariant = "" Then
    MsgBox ("Preencha o campo: (E-mail).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Telefone. ~~ '
If telefoneVariant = "" Then
    MsgBox ("Preencha o campo: (Telefone).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(telefoneVariant) Then
    MsgBox ("Insira apenas números no campo: (Telefone).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Len(telefoneVariant) < 9 Then
    MsgBox ("Insira todos os dígitos no telefone.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Nome. ~~ '
If nomeVariant = "" Then
    MsgBox ("Preencha o campo: (Nome).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Sobrenome. ~~ '
If sobrenomeVariant = "" Then
    MsgBox ("Preencha o campo: (Sobrenome).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Órgão. ~~ '
If orgaoVariant = "" Then
    MsgBox ("Preencha o campo: (Órgão).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Licitação. ~~ '
If licitacaoVariant = "" Then
    MsgBox ("Preencha o campo: (Nº Licitação).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Carona. ~~ '
If caronaVariant = "" Then
    MsgBox ("Preencha o campo: (Carona Governo).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsInArray(caronaVariant, CaronaArr) Then
    MsgBox ("Insira um valor válido no campo: (Carona Governo).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Data Prevista. ~~ '
If dataprevistaVariant = "" Then
    MsgBox ("Preencha o campo: (Data Prevista).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsInArray(dataprevistaVariant, DataPrevistaArr) Then
    MsgBox ("Insira um valor válido no campo: (Data Prevista).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ SKU. ~~ '
If skuVariant = "" Then
    MsgBox ("Preencha o campo: (SKU).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(skuVariant) Then
    MsgBox ("Insira apenas números no campo: (SKU).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Quantidade total. ~~ '
If qtdtotalVariant = "" Then
    MsgBox ("Preencha o campo: (Qtd. Peças (TOTAL)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(qtdtotalVariant) Then
    MsgBox ("Insira apenas números no campo: (Qtd. Peças (TOTAL)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
qtdtotalVariantInt = CInt(qtdtotalVariant)

' ~~ Quantidade faturado. ~~ '
If qtdfaturadoVariant = "" Then
    MsgBox ("Preencha o campo: (Qtd. Peças (FATURADO)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(qtdfaturadoVariant) Then
    MsgBox ("Insira apenas números no campo: (Qtd. Peças (FATURADO)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
qtdfaturadoVariantInt = CInt(qtdfaturadoVariant)

' ~~ Checando se quantidade faturado é maior que quantidade total. ~~ '
If qtdfaturadoVariantInt > qtdtotalVariantInt Then
    MsgBox ("Quantidade faturada não pode ser maior que quantidade total.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Faturamento total. ~~ '
If fattotalVariant = "" Then
    MsgBox ("Preencha o campo: (Valor (TOTAL)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(fattotalVariant) Then
    MsgBox ("Insira apenas números no campo: (Valor (TOTAL)). Pode ser utilizado ponto e vírgula.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
fattotalVariantCur = CCur(fattotalVariant)

' ~~ Faturado. ~~ '
If faturadoVariant = "" Then
    MsgBox ("Preencha o campo: (Valor (FATURADO)).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsNumeric(faturadoVariant) Then
    MsgBox ("Insira apenas números no campo: (Valor (FATURADO)). Pode ser utilizado ponto e vírgula.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
faturadoVariantCur = CCur(faturadoVariant)

' ~~ Checando se faturado total é maior que faturamento total. ~~ '
If faturadoVariantCur > fattotalVariantCur Then
    MsgBox ("Valor faturado não pode ser maior que faturamento total.")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Probabilidade. ~~ '
If probabilidadeVariant = "" Then
    MsgBox ("Preencha o campo: (Probabilidade).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsInArray(probabilidadeVariant, ProbabilidadeArr) Then
    MsgBox ("Insira um valor válido no campo: (Probabilidade).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ Status. ~~ '
If statusVariant = "" Then
    MsgBox ("Preencha o campo: (Status).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If
If Not IsInArray(statusVariant, StatusArr) Then
    MsgBox ("Insira um valor válido no campo: (Status).")
    ActiveSheet.Protect Password:=senha, AllowFiltering:=True
    Exit Sub
End If

' ~~ UCase. ~~ '
emailVariant = UCase(emailVariant)
nomeVariant = UCase(nomeVariant)
sobrenomeVariant = UCase(sobrenomeVariant)
orgaoVariant = UCase(orgaoVariant)
customizadoVariant = UCase(customizadoVariant)
licitacaoVariant = UCase(licitacaoVariant)
caronaVariant = UCase(caronaVariant)
comentariosVariant = UCase(comentariosVariant)

' ~~ Inserindo no PIPE. ~~ '
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "C") = cnpjVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "D") = emailVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "E") = telefoneVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "F") = nomeVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "G") = sobrenomeVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "H") = orgaoVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "I") = dataprevistaVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "J") = skuVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "N") = customizadoVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "O") = qtdtotalVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "P") = qtdfaturadoVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "R") = fattotalVariantCur
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "S") = faturadoVariantCur
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "U") = licitacaoVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "V") = caronaVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "W") = comentariosVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "X") = probabilidadeVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "Y") = statusVariant
ThisWorkbook.Worksheets("PIPE").Cells(linhaModificar, "Z") = numeropedidonfVariant

' ~~ Inserindo no CONSOLIDADO. ~~ '
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "C") = cnpjVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "D") = emailVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "E") = telefoneVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "F") = nomeVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "G") = sobrenomeVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "H") = orgaoVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "I") = dataprevistaVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "J") = skuVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "N") = customizadoVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "O") = qtdtotalVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "P") = qtdfaturadoVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "R") = fattotalVariantCur
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "S") = faturadoVariantCur
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "U") = licitacaoVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "V") = caronaVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "W") = comentariosVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "X") = probabilidadeVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "Y") = statusVariant
ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "Z") = numeropedidonfVariant

' ~~ Checando se os valores consultados são diferentes dos salvos e pintando de verde. ~~ '
If cnpjVariant <> CNPJCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "C").Interior.Color = RGB(0, 255, 0)
End If
If emailVariant <> EMAILCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "D").Interior.Color = RGB(0, 255, 0)
End If
If telefoneVariant <> TELEFONECON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "E").Interior.Color = RGB(0, 255, 0)
End If
If nomeVariant <> NOMECON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "F").Interior.Color = RGB(0, 255, 0)
End If
If sobrenomeVariant <> SOBRENOMECON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "G").Interior.Color = RGB(0, 255, 0)
End If
If orgaoVariant <> ORGAOCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "H").Interior.Color = RGB(0, 255, 0)
End If
If dataprevistaVariant <> DATAPREVISTACON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "I").Interior.Color = RGB(0, 255, 0)
End If
If skuVariant <> SKUCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "J").Interior.Color = RGB(0, 255, 0)
End If
If customizadoVariant <> CUSTOMIZADOCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "N").Interior.Color = RGB(0, 255, 0)
End If
If qtdtotalVariant <> QTDTOTALCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "O").Interior.Color = RGB(0, 255, 0)
End If
If qtdfaturadoVariant <> QTDFATURADOCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "P").Interior.Color = RGB(0, 255, 0)
End If
If fattotalVariant <> FATTOTALCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "R").Interior.Color = RGB(0, 255, 0)
End If
If faturadoVariant <> FATURADOCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "S").Interior.Color = RGB(0, 255, 0)
End If
If licitacaoVariant <> LICITACAOCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "U").Interior.Color = RGB(0, 255, 0)
End If
If caronaVariant <> CARONACON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "V").Interior.Color = RGB(0, 255, 0)
End If
If comentariosVariant <> COMENTARIOSCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "W").Interior.Color = RGB(0, 255, 0)
End If
If probabilidadeVariant <> PROBABILIDADECON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "X").Interior.Color = RGB(0, 255, 0)
End If
If statusVariant <> STATUSCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "Y").Interior.Color = RGB(0, 255, 0)
End If
If numeropedidonfVariant <> NUMEROPEDIDONFCON Then
    ThisWorkbook.Worksheets("CONSOLIDADO").Cells(linhaModificar, "Z").Interior.Color = RGB(0, 255, 0)
End If

' ~~ Ajustando colunas. ~~ '
ThisWorkbook.Worksheets("CONSOLIDADO").Columns.AutoFit
ThisWorkbook.Worksheets("PIPE").Range(ThisWorkbook.Worksheets("PIPE").Cells(3, "C"), ThisWorkbook.Worksheets("PIPE").Cells(5000, "Z")).Columns.AutoFit

' ~~ Verifica se status é Finalizado/Ganho. ~~ '
If STATUS.Value = "Finalizado/Ganho" Then
    If NUMEROPEDIDONF.Value = "" Then
        MsgBox ("Finalizado/Ganho." & vbCrLf & vbCrLf & "Assim que tiver o pedido ou NF, não esqueça de preencher o campo.")
    End If
End If

ActiveSheet.Protect Password:=senha, AllowFiltering:=True
Unload Me

End Sub