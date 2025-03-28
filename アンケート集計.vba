Sub CollectSurveyData()
    Dim MyFolder As String
    Dim MyFile As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim TargetRow As Long
    Dim SourceSheet As Worksheet ' アンケートシートを格納
    
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets("集計用シート") ' 集計用シート名を確認

    MyFolder = "C:\YourFolderPath" '指定のフォルダパス、実際の環境に書き直すこと
    MyFile = Dir(MyFolder & "\*.xlsx") 'フォルダ内の.xlsxファイルを探索
    
    Do While MyFile <> ""
        Set wb = Workbooks.Open(MyFolder & "\" & MyFile)
        Set SourceSheet = wb.Sheets(1) ' 必要であればシート名を指定

        If Not SourceSheet Is Nothing Then
            TargetRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
            
            ' データをコピー（空白チェックを追加）
            If SourceSheet.Range("C2").Value <> "" Then ws.Cells(TargetRow, 1).Value = SourceSheet.Range("C2").Value
            If SourceSheet.Range("C3").Value <> "" Then ws.Cells(TargetRow, 2).Value = SourceSheet.Range("C3").Value
            If SourceSheet.Range("C4").Value <> "" Then ws.Cells(TargetRow, 3).Value = SourceSheet.Range("C4").Value
            If SourceSheet.Range("C5").Value <> "" Then ws.Cells(TargetRow, 4).Value = SourceSheet.Range("C5").Value
            If SourceSheet.Range("C6").Value <> "" Then ws.Cells(TargetRow, 5).Value = SourceSheet.Range("C6").Value
        End If
        
        wb.Close SaveChanges:=False
        MyFile = Dir
    Loop
     'MsgBox "現在処理中のファイル: " & MyFile
     'MsgBox "転記先の行番号: " & TargetRow

    
    MsgBox "集計が完了しました！", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

