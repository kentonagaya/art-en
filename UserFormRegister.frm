Option Explicit

' --- 登録ボタン ---
Private Sub RegisterButton_Click()

    Dim rng As Range
    Dim newran As Long
    Dim rc As VbMsgBoxResult
    With Worksheets("顧客一覧")

        Set rng = .Range("B2:C999").Find(Me.SearchBoxSlip.Value)
        
        If rng Is Nothing Then
        
            Set rng = .Range("B2:C999").Find(Me.SearchBoxSale.Value)
            
            If Not rng Is Nothing Then
                MsgBox "空白または売りカタが重複しています。"
                Exit Sub
            End If
            
        Else
        
            MsgBox "空白または伝票名が重複しています。"
            Exit Sub
        
        End If
        
        newran = .Cells(.Rows.count, "B").End(xlUp).Row + 1
        
        Me.SearchBoxNumber.Value = Format(Me.SearchBoxNumber.Value, "000")

        rc = MsgBox("番号：" & Me.SearchBoxNumber.Value & Chr(13) & "伝票名（屋号）：" & Me.SearchBoxSlip.Text & Chr(13) & "売りカタ：" & Me.SearchBoxSale.Text, vbYesNo + vbQuestion, "登録確認 ")
        
        If rc <> vbYes Then Exit Sub
        
        .Cells(newran, 1).Value = Me.SearchBoxNumber.Text
        .Cells(newran, 2).Value = Me.SearchBoxSlip.Text
        .Cells(newran, 3).Value = Me.SearchBoxSale.Text
        
        MsgBox ("登録完了です。")
        
        Me.SearchBoxSlip.Value = ""
        Me.SearchBoxNumber.Value = ""
        Me.SearchBoxSale.Value = ""
    
    End With

End Sub

' --- 名前と売りカタ同じボタン ---
Private Sub SameValueButton_Click()

    Me.SearchBoxSale.Text = Me.SearchBoxSlip

End Sub

' --- 閉じるボタン ---
Private Sub CloseButton_Click()

    Unload UserFormRegister

End Sub

Private Sub UserForm_Click()

End Sub
