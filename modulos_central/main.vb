' ~~ Função para identificar última linha preenchida em uma coluna. ~~ '
Function lastRow(workbookName As String, worksheetName As String, coluna As String)

lastRow = Workbooks(workbookName).Worksheets(worksheetName).Cells(ThisWorkbook.Worksheets(worksheetName).Rows.Count, coluna).End(xlUp).Row

End Function

' ~~ Ao abrir planilha. ~~ '
Private Sub Workbook_Open()

' ~~ Declarando variáveis. ~~ '
Dim centralWorkbook As Workbook
Dim targetWorkbook As Workbook
Dim sourceRange As Range
Dim targetRange As Range
Dim rng As Range
Set centralWorkbook = ThisWorkbook

' ~~ Limpando CONSOLIDADO CENTRAL. ~~ '
Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A4:Z5000")
targetRange.Interior.Color = RGB(218, 233, 248)
targetRange.ClearContents

' ~~ Desativando macros de abertura. ~~ '
Application.EnableEvents = False

' ~~ Coleta path da planilha CENTRAL. ~~ '
diretorioCentral = ThisWorkbook.Path

' ~~ Copiando CONSOLIDADO do André e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - ANDRE.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - ANDRÉ").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - ANDRÉ").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - ANDRÉ", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - ANDRÉ").Range("A1", ThisWorkbook.Worksheets("PIPE - ANDRÉ").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A3")
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO do Douglas e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - DOUGLAS.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - DOUGLAS").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - DOUGLAS").Range("A:Z")
targetRange.EntireColumn.AutoFit

' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - DOUGLAS", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - DOUGLAS").Range("A2", ThisWorkbook.Worksheets("PIPE - DOUGLAS").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO do França e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - FRANÇA.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - FRANÇA").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - FRANÇA").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - FRANÇA", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - FRANÇA").Range("A2", ThisWorkbook.Worksheets("PIPE - FRANÇA").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO do Gustavo e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - GUSTAVO.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - GUSTAVO").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - GUSTAVO").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - GUSTAVO", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - GUSTAVO").Range("A2", ThisWorkbook.Worksheets("PIPE - GUSTAVO").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO da Luana e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - LUANA.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - LUANA").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - LUANA").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - LUANA", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - LUANA").Range("A2", ThisWorkbook.Worksheets("PIPE - LUANA").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO do Maikon e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - MAIKON.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - MAIKON").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - MAIKON").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - MAIKON", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - MAIKON").Range("A2", ThisWorkbook.Worksheets("PIPE - MAIKON").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO do Marcelo e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - MARCELO.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - MARCELO").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - MARCELO").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - MARCELO", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - MARCELO").Range("A2", ThisWorkbook.Worksheets("PIPE - MARCELO").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO da Margarida e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - MARGARIDA.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - MARGARIDA").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - MARGARIDA").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - MARGARIDA", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - MARGARIDA").Range("A2", ThisWorkbook.Worksheets("PIPE - MARGARIDA").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO da Margarida e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - RAQUEL.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - RAQUEL").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - RAQUEL").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - RAQUEL", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - RAQUEL").Range("A2", ThisWorkbook.Worksheets("PIPE - RAQUEL").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO da Renata e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - RENATA.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - RENATA").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - RENATA").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - RENATA", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - RENATA").Range("A2", ThisWorkbook.Worksheets("PIPE - RENATA").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO do Ronaldo e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - RONALDO.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - RONALDO").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - RONALDO").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - RONALDO", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - RONALDO").Range("A2", ThisWorkbook.Worksheets("PIPE - RONALDO").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Definindo linha atual do CONSOLIDADO para colar próximo range. ~~ '
linhaAtual = lastRow("PIPE REVENDAS - CENTRAL", "CONSOLIDADO", "A") + 1

' ~~ Copiando CONSOLIDADO do Youhanna e importando na CENTRAL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - YOUHANNA.xlsm", ReadOnly:=True)
Set sourceRange = targetWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000")
Set targetRange = centralWorkbook.Worksheets("PIPE - YOUHANNA").Range("A1")
sourceRange.Copy
targetRange.PasteSpecial Paste:=xlPasteValues
targetRange.PasteSpecial Paste:=xlPasteFormats
Set targetRange = centralWorkbook.Worksheets("PIPE - YOUHANNA").Range("A:Z")
targetRange.EntireColumn.AutoFit
' ~~ Colando no CONSOLIDADO CENTRAL. ~~ '
ultimaLinha = lastRow("PIPE REVENDAS - CENTRAL", "PIPE - YOUHANNA", "C")
If Not ultimaLinha = 1 Then
    Set sourceRange = centralWorkbook.Worksheets("PIPE - YOUHANNA").Range("A2", ThisWorkbook.Worksheets("PIPE - YOUHANNA").Cells(ultimaLinha, "Z"))
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A" & linhaAtual)
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll
    Set targetRange = centralWorkbook.Worksheets("CONSOLIDADO").Range("A:Z")
    targetRange.EntireColumn.AutoFit
End If
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=False

' ~~ Ajustando colunas do CONSOLIDADO. ~~ '
ThisWorkbook.Worksheets("CONSOLIDADO").Range("A3:Z5000").Columns.AutoFit
ThisWorkbook.Worksheets("CONSOLIDADO").Range("AB3:AM5000").Columns.AutoFit

' ~~ Ativando macros de abertura. ~~ '
Application.EnableEvents = True

' ~~ Fecha abas abertas. ~~ '
If Worksheets("CONSOLIDADO").Visible = True Then
    Worksheets("CONSOLIDADO").Visible = False
End If
If Worksheets("CONTAS DYNAMICS").Visible = True Then
    Worksheets("CONTAS DYNAMICS").Visible = False
End If
If Worksheets("PRODUTOS").Visible = True Then
    Worksheets("PRODUTOS").Visible = False
End If
If Worksheets("PIPE - ANDRÉ").Visible = True Then
    Worksheets("PIPE - ANDRÉ").Visible = False
End If
If Worksheets("PIPE - DOUGLAS").Visible = True Then
    Worksheets("PIPE - DOUGLAS").Visible = False
End If
If Worksheets("PIPE - GUSTAVO").Visible = True Then
    Worksheets("PIPE - GUSTAVO").Visible = False
End If
If Worksheets("PIPE - MAIKON").Visible = True Then
    Worksheets("PIPE - MAIKON").Visible = False
End If
If Worksheets("PIPE - FRANÇA").Visible = True Then
    Worksheets("PIPE - FRANÇA").Visible = False
End If
If Worksheets("PIPE - RENATA").Visible = True Then
    Worksheets("PIPE - RENATA").Visible = False
End If
If Worksheets("PIPE - MARCELO").Visible = True Then
    Worksheets("PIPE - MARCELO").Visible = False
End If
If Worksheets("PIPE - RONALDO").Visible = True Then
    Worksheets("PIPE - RONALDO").Visible = False
End If
If Worksheets("PIPE - LUANA").Visible = True Then
    Worksheets("PIPE - LUANA").Visible = False
End If
If Worksheets("PIPE - YOUHANNA").Visible = True Then
    Worksheets("PIPE - YOUHANNA").Visible = False
End If
If Worksheets("PIPE - RAQUEL").Visible = True Then
    Worksheets("PIPE - RAQUEL").Visible = False
End If
If Worksheets("PIPE - MARGARIDA").Visible = True Then
    Worksheets("PIPE - MARGARIDA").Visible = False
End If

' ~~ Encerramento. ~~ '
End Sub