VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CopiarNomesArquivos()
    Dim FSO As Object
    Dim folderPath As String
    Dim file As Object
    Dim nextRow As Long
    Dim dialogBox As FileDialog
    Dim erroLog As String
    
    ' Abre o diálogo para selecionar a pasta
    Set dialogBox = Application.FileDialog(msoFileDialogFolderPicker)
    dialogBox.Title = "Selecione a Pasta Contendo os Arquivos Excel"
    
    ' Verifica se um diretório foi selecionado
    If dialogBox.Show = -1 Then
        folderPath = dialogBox.SelectedItems(1)
    Else
        MsgBox "Nenhum diretório selecionado. Operação cancelada.", vbExclamation
        Exit Sub
    End If
    
    ' Cria o objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verifica se o diretório especificado existe
    If Not FSO.FolderExists(folderPath) Then
        MsgBox "O diretório especificado não existe.", vbCritical
        Exit Sub
    End If
    
    ' Define a linha inicial na planilha "Pasta2"
    nextRow = ThisWorkbook.Sheets("Pasta2").Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    ' Percorre todos os arquivos na pasta
    For Each file In FSO.GetFolder(folderPath).Files
        If LCase(file.Name) Like "*.xls*" Then
            ' Copia o nome do arquivo para a planilha "Pasta2"
            ThisWorkbook.Sheets("Pasta2").Cells(nextRow, 1).Value = file.Name
            nextRow = nextRow + 1
        End If
    Next file
    
    ' Libera objetos
    Set FSO = Nothing
    Set dialogBox = Nothing
    
    ' Exibe mensagem de conclusão
    MsgBox "Nomes dos arquivos copiados com sucesso.", vbInformation
End Sub

