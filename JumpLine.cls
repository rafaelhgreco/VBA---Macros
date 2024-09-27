VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub InserirJustificativa()
    Dim i As Long
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row ' Substitua "A" pela coluna que contém seus dados principais
    For i = lastRow To 1 Step -1
        Rows(i + 1).Insert Shift:=xlDown
        ' Preencha a linha inserida com "justificativa" ou outros dados desejados
        Cells(i + 1, "A").Value = "Justificativa" ' Substitua "A" pela coluna desejada
        ' Aplique formatações adicionais se necessário
        ' Exemplo: Cells(i + 1, "A").Font.Bold = True
    Next i
End Sub

