VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CopiarParaPasta2()
    Dim FSO As Object
    Dim folderPath As String
    Dim file As Object
    Dim wb As Workbook
    Dim wsDadosBasicos As Worksheet
    Dim nextRow As Long
    Dim lastRow As Long
    Dim dialogBox As FileDialog
    Dim erroLog As String
    Dim i As Long

    Set dialogBox = Application.FileDialog(msoFileDialogFolderPicker)
    dialogBox.Title = "Selecione a Pasta Contendo os Arquivos Excel"
    
    If dialogBox.Show = -1 Then
        folderPath = dialogBox.SelectedItems(1)
    Else
        MsgBox "Nenhum diretório selecionado. Operação cancelada.", vbExclamation
        Exit Sub
    End If
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Not FSO.FolderExists(folderPath) Then
        MsgBox "O diretório especificado não existe.", vbCritical
        Exit Sub
    End If
    
    erroLog = ""

    For Each file In FSO.GetFolder(folderPath).Files
        If LCase(file.Name) Like "*.xls*" Then
            Set wb = Workbooks.Open(file.Path)
            
            On Error Resume Next
            Set wsDadosBasicos = wb.Sheets("Dados básicos")    'ajustar abas
            On Error GoTo 0
            
            If Not wsDadosBasicos Is Nothing Then
                lastRow = ThisWorkbook.Sheets("Pasta2").Cells(Rows.Count, 1).End(xlUp).Row
                
                With ThisWorkbook.Sheets("Pasta2")
                    ' Loop através do intervalo B9:B100, verificando e copiando apenas células não vazias
                    For i = 9 To 100
                        If Trim(wsDadosBasicos.Cells(i, 2).Value) <> "" Then ' Verifica se a célula não está vazia
                            lastRow = lastRow + 1
                            .Cells(lastRow, 1).Value = wsDadosBasicos.Cells(i, 2).Value ' Coluna B
                            .Cells(lastRow, 2).Value = wsDadosBasicos.Cells(i, 5).Value ' Coluna E
                            .Cells(lastRow, 3).Value = wsDadosBasicos.Cells(i, 7).Value ' Coluna G
                            .Cells(lastRow, 4).Value = wsDadosBasicos.Cells(i, 9).Value ' Coluna I
                            .Cells(lastRow, 5).Value = wsDadosBasicos.Cells(i, 79).Value ' Coluna CA
                            .Cells(lastRow, 6).Value = file.Name ' Nome do arquivo
                        End If
                    Next i
                End With
                
                wb.Close False
            Else
                erroLog = erroLog & vbCrLf & "A aba 'Dados básicos' não foi encontrada no arquivo " & file.Name
            End If
            
            Set wsDadosBasicos = Nothing
        End If
    Next file
    
    If Len(erroLog) > 0 Then
        MsgBox "Os seguintes erros foram encontrados:" & vbCrLf & erroLog, vbCritical
    End If
    
    Set FSO = Nothing
End Sub



