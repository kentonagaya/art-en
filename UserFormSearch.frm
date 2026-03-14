Option Explicit

' 親フォーム参照
Public ParentForm As UserFormImport
' 検索モード（SELLER / BUYER）
Public SearchMode As String

' --- ユーザーフォーム初期化---
Private Sub UserForm_Activate()

    Select Case UCase(SearchMode)
        Case "SELLER"
            Me.Caption = "売主検索"
            Me.ListBox1.BackColor = &HFFFFC0
            Me.SelectButton.BackColor = &HFFFF80
        Case "BUYER"
            Me.Caption = "買主検索"
            Me.ListBox1.BackColor = &HC0FFC0
            Me.SelectButton.BackColor = &H80FF80
        Case Else
            Me.Caption = "検索フォーム"
    End Select

End Sub

' --- リスト初期化 ---
Private Sub UserForm_Initialize()
    
    LoadParticipantList

End Sub


' --- 検索窓クリアボタン ---
Private Sub ClearButton_Click()

    Me.SearchBox.Value = ""
    LoadParticipantList

End Sub

' --- 検索ボタン ---
Private Sub SearchButton_Click()

    Dim keyword As String
    keyword = Trim(Me.SearchBox.Value)

    LoadParticipantList keyword

    If Me.ListBox1.ListCount = 0 Then
        MsgBox "該当者が見つかりませんでした。", vbInformation
    End If

End Sub

' --- 参加者リスト読み込み・検索対応 ---
Private Sub LoadParticipantList(Optional ByVal keyword As String = "")

    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim clientNameStr As String, shortNameStr As String, clientNumberStr As String

    Set ws = Worksheets("項目")

    Me.ListBox1.Clear

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    With Me.ListBox1
        .Clear
        .ColumnCount = 3
        .ColumnWidths = "40;120;120"
    End With

    For i = 2 To lastRow

        clientNameStr = CStr(ws.Cells(i, 2).Value)   ' 屋号
        shortNameStr = CStr(ws.Cells(i, 3).Value)   ' 売りカタ
        clientNumberStr = CStr(ws.Cells(i, 1).Value) ' 参加者番号

        ' keywordなし＝全件
        ' keywordあり＝部分一致
        If keyword = "" _
           Or InStr(clientNameStr, keyword) > 0 _
           Or InStr(shortNameStr, keyword) > 0 _
           Or InStr(clientNumberStr, keyword) > 0 Then

            Me.ListBox1.AddItem ws.Cells(i, 1).Value ' 3列目：参加者番号
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = ws.Cells(i, 2).Value ' 1列目：屋号
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = ws.Cells(i, 3).Value ' 2列目：売りカタ

        End If

    Next i

End Sub

' --- 選択ボタン ---
Private Sub SelectButton_Click()
    
    Dim clientName As String, shortName As String, clientNumber As String
    
    ' 未入力チェック
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "出力する対象を選択してください。", vbExclamation
        Exit Sub
    End If

    ' データの取得
    clientNumber = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
    clientName = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
    shortName = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
    
    ' 親フォームへ値を渡す
    Select Case UCase(SearchMode)
        Case "SELLER"
            ParentForm.売主.Text = clientNumber
            ParentForm.うりな.Text = clientName
        Case "BUYER"
            ParentForm.買主.Text = clientNumber
            ParentForm.かいな.Text = shortName
        Case Else
            MsgBox "検索モードが不正です。", vbCritical
    End Select

    Unload Me
    
End Sub

' --- 閉じるボタン ---
Private Sub CloseButton_Click()

    Unload Me
    
End Sub


