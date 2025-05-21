' ~~ Função para remover acentos. ~~ '
Function Acento(Caract As String)

Dim A As String
Dim B As String
Dim i As Integer
Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"

For i = 1 To Len(AccChars)
    A = Mid(AccChars, i, 1)
    B = Mid(RegChars, i, 1)
    Caract = Replace(Caract, A, B)
Next

Acento = Caract

End Function

' ~~ Consulta de conta. ~~ '
Sub consultar_conta()

' ~~ Variáveis ~~ '
Dim requisicao As New MSXML2.XMLHTTP60
Dim resposta As Object
Dim url As String
Dim cnpjconsultado As String
Dim payload As New Dictionary
Dim semAcento As String

' ~~ Informando valores. ~~ '
linhaConsulta = ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(2, "A").Value
cnpjconsultado = ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "A").Value
url = "https://publica.cnpj.ws/cnpj/" & cnpjconsultado

' ~~ Realizando requisição Receita. ~~ '
requisicao.Open "Get", url, False
requisicao.Send

' ~~ Se ocorrer erro na requisição. ~~ '
If requisicao.Status <> 200 Then
    MsgBox "Erro ao consultar: " & requisicao.responseText
    Exit Sub
End If

' ~~ Convertendo para JSON a resposta da requisição. ~~ '
Set resposta = JsonConverter.ParseJson(requisicao.responseText)

' ~~ Integrando dados da Receita na planilha ~~ '
ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "B") = resposta("razao_social")
ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "C") = resposta("estabelecimento")("tipo_logradouro") & " " & resposta("estabelecimento")("logradouro") & ", " & resposta("estabelecimento")("numero") & " " & resposta("estabelecimento")("complemento")
ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "D") = resposta("estabelecimento")("bairro")
ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "E") = resposta("estabelecimento")("cep")
ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "F") = resposta("estabelecimento")("cidade")("nome")
ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "G") = resposta("estabelecimento")("estado")("sigla")

' ~~ Tirando acento. ~~ '
semAcento = ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "C").Value
semAcento = Acento(semAcento)
ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "C") = semAcento

semAcento = ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "D").Value
semAcento = Acento(semAcento)
ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "D") = semAcento

semAcento = ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "F").Value
semAcento = Acento(semAcento)
ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "F") = semAcento

' ~~ UCase ~~ '
Dim rng As Range
Set rngInicial = ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "B")
Set rngFinal = ThisWorkbook.Worksheets("CONTAS DYNAMICS").Cells(linhaConsulta, "G")
For Each rng In Range(rngInicial, rngFinal)
    rng.Value = UCase(rng.Value)
Next rng

' ~~ Fim do código. ~~ '
End Sub