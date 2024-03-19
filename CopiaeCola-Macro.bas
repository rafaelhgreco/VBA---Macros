Attribute VB_Name = "Módulo2"
Sub CopiarParaPasta2()
    Dim FSO As Object
    Dim folderPath As String
    Dim file As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nextRow As Long
    
    ' Definir o caminho do diretório
    folderPath = " "
    
    ' Inicializar o objeto FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar se o diretório existe
    If Not FSO.FolderExists(folderPath) Then
        MsgBox "O diretório especificado não existe.", vbCritical
        Exit Sub
    End If
    
    ' Definir a próxima linha na planilha "Pasta1"
    nextRow = 1
    
    ' Iterar sobre os arquivos no diretório
    For Each file In FSO.GetFolder(folderPath).Files
        ' Verificar se o arquivo é um arquivo Excel
        If InStr(file.Name, ".xlsx") > 0 Or InStr(file.Name, ".xls") > 0 Then
            ' Abrir o arquivo
            Set wb = Workbooks.Open(file.Path)
            
            ' Verificar se a planilha existe
            On Error Resume Next
            Set ws = wb.Sheets("Dados básicos")
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                ' Copiar o conteúdo da célula G9
                ws.Range("G9:G100").Copy
                
                ' Determinar a próxima linha na Pasta1
                nextRow = ThisWorkbook.Sheets("Pasta2").Cells(Rows.Count, "A").End(xlUp).Row + 1
                
                ' Colar na próxima linha na Pasta1
                ThisWorkbook.Sheets("Pasta2").Cells(nextRow, "A").PasteSpecial Paste:=xlPasteValues
                
                ' Escrever o nome do arquivo na próxima coluna
                ThisWorkbook.Sheets("Pasta2").Cells(nextRow, 2).Value = file.Name
                
                ' Incrementar para a próxima linha
                nextRow = nextRow + 1
                
                ' Fechar o arquivo sem salvar alterações
                wb.Close False
            Else
                MsgBox "A planilha 'Dados básicos' não foi encontrada no arquivo " & file.Name, vbExclamation
            End If
        End If
    Next file
    
    ' Liberar o objeto FileSystemObject
    Set FSO = Nothing
End Sub


