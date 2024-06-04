VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
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
    Dim wsDadosCentro As Worksheet
    Dim nextRow As Long
    Dim lastRow As Long
    Dim dialogBox As FileDialog
    Dim erroLog As String
    
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
    
    nextRow = 1
    
    For Each file In FSO.GetFolder(folderPath).Files
        If LCase(file.Name) Like "*.xls*" Then
            Set wb = Workbooks.Open(file.Path)
            
            On Error Resume Next
            Set wsDadosBasicos = wb.Sheets("Dados básicos")
            Set wsDadosCentro = wb.Sheets("Dados de centro")
            On Error GoTo 0
            
            If Not wsDadosBasicos Is Nothing And Not wsDadosCentro Is Nothing Then
                lastRow = ThisWorkbook.Sheets("Pasta2").Cells(Rows.Count, 1).End(xlUp).Row
                ThisWorkbook.Sheets("Pasta2").Cells(lastRow + 1, 1).Resize(92, 1).Value = wsDadosBasicos.Range("G9:G100").Value
                ThisWorkbook.Sheets("Pasta2").Cells(lastRow + 1, 2).Resize(92, 1).Value = wsDadosCentro.Range("EL9:EL100").Value
                ThisWorkbook.Sheets("Pasta2").Cells(lastRow + 1, 3).Value = file.Name
                
                wb.Close False
            Else
                erroLog = erroLog & vbCrLf & "As abas 'Dados básicos' e 'Dados de centro' não foram encontradas no arquivo " & file.Name
            End If
            
            Set wsDadosBasicos = Nothing
            Set wsDadosCentro = Nothing
        End If
    Next file
    
    If Len(erroLog) > 0 Then
        MsgBox "Os seguintes erros foram encontrados:" & vbCrLf & erroLog, vbCritical
    End If
    
    Set FSO = Nothing
End Sub

