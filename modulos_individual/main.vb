' ~~ Variáveis públicas. ~~ '
Public lineValue As Variant
Public listaFinalizados As String
Public linha As Integer

' ~~ Função para encontrar última linha. ~~ '
Function lastRow(coluna As String)

lastRow = ThisWorkbook.Worksheets("PIPE").Cells(ThisWorkbook.Worksheets("PIPE").Rows.Count, coluna).End(xlUp).Row

End Function

' ~~ Ao abrir planilha. ~~ '
Public Sub Workbook_Open()

' ~~ Verifica linhas que estão como Finalizado/Ganho e mostra um aviso. ~~ '
ultimaLinha = lastRow("C") + 1
listaFinalizados = ""
For linha = 1 To ultimaLinha
    If ThisWorkbook.Worksheets("PIPE").Cells(linha, "Y") = "Finalizado/Ganho" Then ' <= Se for Finalizado/Ganho. '
        If ThisWorkbook.Worksheets("PIPE").Cells(linha, "Z") = "" Then ' <= Se Pedido/NF estiver vazio. '
            If listaFinalizados = "" Then ' <= Checa se é a primeira linha sendo inserida. '
                listaFinalizados = CStr(linha)
            Else
                listaFinalizados = listaFinalizados & ", " & CStr(linha)
            End If
            If ThisWorkbook.Worksheets("PIPE").Cells(linha, "C") = "" Then ' <= Quando não encontra mais linhas. '
                listaFinalizados = listaFinalizados & "."
            End If
        End If
    End If
Next

' ~~ Mostra MsgBox com os Pedidos/NF em branco. ~~ '
If Not listaFinalizados = "" Then
    MsgBox ("Lista dos registros finalizados que ainda estão sem Pedido/NF:" & vbCrLf & vbCrLf & listaFinalizados & vbCrLf & vbCrLf & "Lembre-se de inserir essa informação assim que tiver.")
End If
listaFinalizados = ""

' ~~ Encerrando algoritmo de abertura. ~~ '
End Sub
