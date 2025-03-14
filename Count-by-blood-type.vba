Sub CountBloodTypes()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long, bloodType As String
    
    ' シートを指定（アクティブシートを使用）
    Set ws = ActiveSheet
    
    ' 最終行を取得（H列の最終行を判定）
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' Dictionary を作成して血液型のカウント
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 初期化（血液型ごとのカウント）
    dict.Add "A", 0
    dict.Add "B", 0
    dict.Add "O", 0
    dict.Add "AB", 0
    
    ' H列のデータをループ処理
    For i = 2 To lastRow ' 1行目がヘッダーの場合、2行目から
        bloodType = Trim(UCase(ws.Cells(i, "H").Value)) ' 大文字に統一し、余分な空白を除去
        If dict.exists(bloodType) Then
            dict(bloodType) = dict(bloodType) + 1
        End If
    Next i
    
    ' 結果をメッセージボックスで表示
    MsgBox "血液型ごとの人数：" & vbCrLf & _
           "A型: " & dict("A") & "人" & vbCrLf & _
           "B型: " & dict("B") & "人" & vbCrLf & _
           "O型: " & dict("O") & "人" & vbCrLf & _
           "AB型: " & dict("AB") & "人", vbInformation, "血液型集計結果"
    
    ' オブジェクト解放
    Set dict = Nothing
End Sub
