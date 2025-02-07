# Excel-VBA-Introduction

## 生年月日から今日時点での年齢を求めるVBA

''
Sub 計算_年齢_FG列()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim birthDate As Date
    Dim todayDate As Date
    Dim age As Integer

    ' シートを指定（アクティブシートを使用）
    Set ws = ActiveSheet

    ' 今日の日付を取得
    todayDate = Date

    ' 最終行を取得（F列の最終行を基準）
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row ' F列は6番目の列

    ' 2行目から最終行までループ（1行目はヘッダーと仮定）
    For i = 2 To lastRow
        ' F列（生年月日）の値を取得
        If IsDate(ws.Cells(i, 6).Value) Then
            birthDate = ws.Cells(i, 6).Value

            ' 年齢を計算（誕生日が来ていない場合は1歳引く）
            age = Year(todayDate) - Year(birthDate)
            If Month(todayDate) < Month(birthDate) Or _
               (Month(todayDate) = Month(birthDate) And Day(todayDate) < Day(birthDate)) Then
                age = age - 1
            End If

            ' G列（年齢）に出力
            ws.Cells(i, 7).Value = age
        Else
            ' 生年月日が無効な場合はG列を空白にする
            ws.Cells(i, 7).Value = ""
        End If
    Next i

    ' 終了メッセージ
    MsgBox "年齢計算が完了しました！", vbInformation
End Sub
''
