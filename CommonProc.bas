
'データシート名
Public Const DATA_SHEET_NAME = "Data"
Public Const SAVE_BOOK_NAME = "Excelデータ"
Public Const SAVE_TEXT_NAME = "データ"
Public Const FAILED = "失敗しました"
Public Const CONFIRM = "確認"


'【概要】Excelファイル群をExcelファイルとして出力する
Public Function ExcelFilesToExcelFile(ByVal arrExcelFilePaths As Variant, _
                                ByVal strPastedSheetName As String) As Boolean
On Error GoTo ExcelFilesToExcelFile_Err

    ExcelFilesToExcelFile = False

    Dim lngLastRow As Long
    Dim lngArrIdx As Long
    Dim lngPasteRow As Long
    Dim lngPastedRow As Long
    Dim objPasteWb As Excel.Workbook
    Dim objPastedWb As Excel.Workbook
    Dim objWs As Excel.Worksheet

    'Excelファイル作成
    Set objPastedWb = Workbooks.Add
    
    'シート名をDATAにする
    ActiveSheet.Name = strPastedSheetName
    
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrExcelFilePaths)
        'Excelファイルを開く
        Workbooks.Open arrExcelFilePaths(lngArrIdx)
        Set objPasteWb = ActiveWorkbook
        '貼り付け先行初期化
        lngPastedRow = 1
        'シートの数だけ繰り返す
        For Each objWs In objPasteWb.Worksheets
            With objPasteWb.Worksheets(objWs.Name)
                '貼り付け元最終行を取得
                lngLastRow = .Cells(1, 1).End(xlDown).Row
                '最終行まで繰り返す
                For lngPasteRow = 1 To lngLastRow
                    '貼り付け元→貼り付け先
                    objPastedWb.Worksheets(strPastedSheetName).Cells(lngPastedRow, 1).Value = .Cells(lngPasteRow, 1).Value
                    '貼り付け先行をカウントアップ
                    lngPastedRow = lngPastedRow + 1
                Next lngPasteRow
            End With
        Next objWs
    Next lngArrIdx
    
    '保存
    objPastedWb.SaveAs SAVE_BOOK_NAME
    '閉じる
    objPasteWb.Close
    objPastedWb.Close
    
    ExcelFilesToExcelFile = True
    
ExcelFilesToExcelFile_Err:

ExcelFilesToExcelFile_Exit:
    Set objPasteWb = Nothing
    Set objPastedWb = Nothing
    Set objWs = Nothing
End Function


'【概要】特定のディレクトリの特定の拡張子のファイルを取得する
Public Function GetFilePaths(ByVal strDirectoryPath As String, _
                        ByVal strExtensionName As String) As Variant
On Error GoTo GetFilePaths_Err
    
    Dim lngArrIdx As Long
    Dim arrRet() As Variant
    Dim objFso As FileSystemObject
    Dim objFile As File
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    With objFso
        'フォルダ内のファイルの数だけ繰り返す
        For Each objFile In .GetFolder(strDirectoryPath).Files
            '拡張子が指定のものだった場合
            If .GetExtensionName(objFile.Name) = strExtensionName Then
                '配列再宣言
                ReDim Preserve arrRet(lngArrIdx)
                '配列格納
                arrRet(lngArrIdx) = objFile.Path
                '配列の要素番号を1つ進める
                lngArrIdx = lngArrIdx + 1
            End If
        Next objFile
    End With
    
    GetFilePaths = arrRet
    
GetFilePaths_Err:

GetFilePaths_Exit:
    Set objFso = Nothing
    Set objFile = Nothing
End Function
