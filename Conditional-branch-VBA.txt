Sub IfExample() ' サブルーチンの開始
　　Dim score As Integer ' 変数scoreを整数型として宣言
　　score = Range("A1").Value ' セルA1の値をscoreに代入
　　If score >= 60 Then ' scoreが60以上の場合の条件分岐を開始
　　　　MsgBox "合格" ' 条件が真の場合、「合格」というメッセージボックスを表示
　　Else ' scoreが60未満の場合の処理
　　　　MsgBox "不合格" ' 条件が偽の場合、「不合格」というメッセージボックスを表示
　　End If ' 条件分岐の終了
End Sub ' サブルーチンの終了
