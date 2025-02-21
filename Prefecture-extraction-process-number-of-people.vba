Sub ExtractAndCountPrefecture()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim rng As Range, cell As Range
    Dim prefectures As Variant
    Dim i As Integer, lastRow As Long
    Dim dict As Object
    
    ' ワークシート設定
    Set ws1 = ThisWorkbook.Sheets("Sheet1") ' データがあるシート
    Set ws2 = ThisWorkbook.Sheets("Sheet2") ' 集計結果を出力するシート
    Set dict = CreateObject("Scripting.Dictionary") ' 都道府県の出現回数を記録

    ' 47都道府県のリスト
    prefectures = Array("北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", _
                        "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", _
                        "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県", _
                        "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", _
                        "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県", _
                        "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県", _
                        "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県")

    ' 最終行を取得
    lastRow = ws1.Cells(ws1.Rows.Count, "M").End(xlUp).Row
    Set rng = ws1.Range("M2:M" & lastRow) ' M列のデータ範囲

    ' 都道府県を抽出
    For Each cell In rng
        For i = LBound(prefectures) To UBound(prefectures)
            If InStr(cell.Value, prefectures(i)) > 0 Then
                cell.Offset(0, 1).Value = prefectures(i) ' N列に都道府県を出力
                ' 集計用Dictionaryに追加
                If dict.exists(prefectures(i)) Then
                    dict(prefectures(i)) = dict(prefectures(i)) + 1
                Else
                    dict(prefectures(i)) = 1
                End If
                Exit For
            End If
        Next i
    Next cell

    ' シート2のA列とB列に集計結果を出力（クリア後に書き込む）
    ws2.Cells.Clear
    ws2.Range("A1").Value = "都道府県"
    ws2.Range("B1").Value = "人数"

    i = 2 ' A2から開始
    Dim key As Variant
    For Each key In dict.keys
        ws2.Cells(i, 1).Value = key ' 都道府県名
        ws2.Cells(i, 2).Value = dict(key) ' 人数
        i = i + 1
    Next key

    ' 列幅を自動調整
    ws2.Columns("A:B").AutoFit

    MsgBox "都道府県の抽出と集計が完了しました！", vbInformation
End Sub


