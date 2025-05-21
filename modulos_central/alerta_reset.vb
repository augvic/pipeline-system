' ~~ Quando clicar em não, fecha tela. ~~ '
Private Sub Não_Click()

Unload Me

End Sub

' ~~ Quando clicar em sim, reseta todas as alterações. ~~ '
Private Sub Sim_Click()

Unload Me

' ~~ Declarando variáveis. ~~ '
Dim rngSource As Range
Dim targetWorkbook As Workbook

' ~~ Resetando CONSOLIDADO CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)

' ~~ Desativando macros de abertura. ~~ '
Application.EnableEvents = False

' ~~ Coleta path da planilha CENTRAL. ~~ '
diretorioCentral = ThisWorkbook.Path

' ~~ Resetando André CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - ANDRÉ").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha André. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - ANDRE.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Douglas CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - DOUGLAS").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha Douglas. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - DOUGLAS.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando França CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - FRANÇA").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha FRANÇA. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - FRANÇA.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Gustavo CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - GUSTAVO").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha GUSTAVO. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - GUSTAVO.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Luana CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - LUANA").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha LUANA. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - LUANA.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Maikon CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - MAIKON").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha MAIKON. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - MAIKON.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Marcelo CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - MARCELO").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha MARCELO. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - MARCELO.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Margarida CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - MARGARIDA").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha MARGARIDA. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - MARGARIDA.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Raquel CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - RAQUEL").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha RAQUEL. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - RAQUEL.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Renata CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - RENATA").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha RENATA. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - RENATA.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Ronaldo CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - RONALDO").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha RONALDO. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - RONALDO.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Resetando Youhanna CENTRAL. ~~ '
Set rngSource = ThisWorkbook.Worksheets("PIPE - YOUHANNA").Range("C2:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
' ~~ Resetando planilha YOUHANNA. ~~ '
Set targetWorkbook = Workbooks.Open(diretorioCentral & "\PIPE - YOUHANNA.xlsm", ReadOnly:=False)
Set rngSource = targetWorkbook.Worksheets("CONSOLIDADO").Range("C4:Z5000")
rngSource.Interior.Color = RGB(218, 233, 248)
targetWorkbook.Close SaveChanges:=True

' ~~ Ativando macros de abertura. ~~ '
Application.EnableEvents = True

' ~~ Encerramento. ~~ '
End Sub