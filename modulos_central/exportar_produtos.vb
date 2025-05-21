' ~~ Ao clicar no botão de exportar produtos. ~~ '
Sub export_products()

' ~~ Declarando variáveis. ~~ '
Dim rngSource As Range
Dim rngTarget As Range
Dim targetWorkbook As Workbook

' ~~ Desativando macros de abertura. ~~ '
Application.EnableEvents = False

' ~~ Importando produtos atualizados no PIPE do André. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - ANDRE.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE do Douglas. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - DOUGLAS.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE do França. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - FRANÇA.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE do Gustavo. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - GUSTAVO.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE da Luana. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - LUANA.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE do Maikon. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - MAIKON.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE do Marcelo. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - MARCELO.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE da Margarida. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - MARGARIDA.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE da Raquel. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - RAQUEL.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE da Renata. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - RENATA.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE do Ronaldo. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - RONALDO.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Importando produtos atualizados no PIPE do Youhanna. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - YOUHANNA.xlsm", ReadOnly:=False)
Set rngSource = ThisWorkbook.Worksheets("PRODUTOS").Range("A4:I5000")
Set rngTarget = targetWorkbook.Worksheets("PRODUTOS").Range("A2:I5000")
rngSource.Copy
rngTarget.PasteSpecial Paste:=xlPasteAll
rngTarget.EntireColumn.AutoFit
Application.CutCopyMode = False
targetWorkbook.Close SaveChanges:=True

' ~~ Ativando macros de abertura. ~~ '
Application.EnableEvents = True

' ~~ Encerrando código. ~~ '
End Sub