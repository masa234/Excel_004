
 
Public Sub 正方形長方形1_Click()
On Error GoTo 正方形長方形1_Click_Err
    
    Dim strOutputExcelFilePath As String
    Dim strTextFileOutputPath As String
    
    '画面の更新をオフにする
    Application.ScreenUpdating = False
    
    'Excelファイルまとめ
    If CombineExcelFiles(DATA_SHEET_NAME) = False Then
        Call MsgBox(FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
    'Excelファイルパス
    strOutputExcelFilePath = ThisWorkbook.Path & "\" & SAVE_BOOK_NAME & ".xlsx"
    
    'テキストファイル出力先
    strTextFileOutputPath = ThisWorkbook.Path & "\" & SAVE_TEXT_NAME & ".txt"
    
    If ExcelFileToTextFile(strOutputExcelFilePath, strTextFileOutputPath) = False Then
        Call MsgBox(FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
    
正方形長方形1_Click_Err:

正方形長方形1_Click_Exit:
    '画面の更新をオンにする
    Application.ScreenUpdating = True
End Sub
