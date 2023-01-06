

'【概要】複数のExcelファイルを1つのExcelファイルにまとめる
Public Function CombineExcelFiles(ByVal strSheetName As String) As Boolean
On Error GoTo CombineExcelFiles_Err

    CombineExcelFiles = False

    Dim arrExcelFilePaths As Variant
    
    'Excelファイル群を取得
    arrExcelFilePaths = GetFilePaths(ThisWorkbook.Path, "xlsm")
    
    '1つのExcelファイルとして出力する
    If ExcelFilesToExcelFile(arrExcelFilePaths, strSheetName) = False Then
        GoTo CombineExcelFiles_Exit
    End If
    
    CombineExcelFiles = True
    
CombineExcelFiles_Err:

CombineExcelFiles_Exit:

End Function


'【概要】Excelファイルをテキストとして出力
Public Function ExcelFileToTextFile(ByVal strOutputExcelFilePath As String, _
                                ByVal strTextFileOutputPath As String) As Boolean
On Error GoTo ExcelFileToTextFile_Err

    ExcelFileToTextFile = False
    
    Dim lngFreeFile As Long
    Dim lngLastRow As Long
    Dim lngCurrentRow As Long
    Dim objWb As Excel.Workbook
    
    'フリーファイルを取得
    lngFreeFile = FreeFile

    'テキストファイルを書き出す
    Open strTextFileOutputPath & ".txt" For Output As #lngFreeFile
    
    'Excelファイルを開く
    Workbooks.Open strOutputExcelFilePath
    Set objWb = ActiveWorkbook
    
    '最終行を取得
    lngLastRow = objWb.Worksheets(ActiveSheet.Name).Cells(1, 1).End(xlDown).Row
    
    '最終行まで繰り返す
    For lngCurrentRow = 1 To lngLastRow
        'ファイル書き出し
        Print #lngFreeFile, Cells(lngCurrentRow, 1).Value
        '行をカウントアップ
        lngCurrentRow = lngCurrentRow + 1
    Next lngCurrentRow
    
    '閉じる
    objWb.Close
        
    ExcelFileToTextFile = True
    
ExcelFileToTextFile_Err:

ExcelFileToTextFile_Exit:
    'テキストファイルを閉じる
    Set objWb = Nothing
End Function


