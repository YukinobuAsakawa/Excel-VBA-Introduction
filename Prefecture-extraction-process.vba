Sub ExtractPrefecture()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim prefectures As Variant
    Dim i As Integer

    ' ワークシート設定
    Set ws = ActiveSheet
    Set rng = ws.Range("M2:M502") ' 都道府県を抽出したい範囲

    ' 47都道府県のリスト
    prefectures = Array("北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", _
                        "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", _
                        "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県", _
                        "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", _
                        "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県", _
                        "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県", _
                        "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県")

    ' データを走査
    For Each cell In rng
        For i = LBound(prefectures) To UBound(prefectures)
            If InStr(cell.Value, prefectures(i)) > 0 Then
                cell.Offset(0, 1).Value = prefectures(i) ' B列に抽出結果を出力
                Exit For
            End If
        Next i
    Next cell
End Sub
