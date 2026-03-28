Option Explicit

Private Const COL_BUYER  As Long = 4
Private Const COL_SELLER As Long = 6

Private matchCells() As String
Private matchCount    As Long
Private currentIndex  As Long

' --- フォーム初期化 ---
Private Sub UserForm_Initialize()
    Me.OptionBuyer.Value = True
    Me.NavLabel.Caption = ""
    LoadParticipantList
End Sub

' --- 参加者リスト読み込み ---
Private Sub LoadParticipantList(Optional ByVal keyword As String = "")
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim numStr As String, nameStr As String, shortStr As String

    Set ws = Worksheets("顧客一覧")

    With Me.ListBox1
        .Clear
        .ColumnCount = 3
        .ColumnWidths = "40;120;120"
    End With

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    For i = 2 To lastRow
        numStr = CStr(ws.Cells(i, 1).Value)
        nameStr = CStr(ws.Cells(i, 2).Value)
        shortStr = CStr(ws.Cells(i, 3).Value)

        If keyword = "" _
           Or InStr(numStr, keyword) > 0 _
           Or InStr(nameStr, keyword) > 0 _
           Or InStr(shortStr, keyword) > 0 Then
            Me.ListBox1.AddItem numStr
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = nameStr
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = shortStr
        End If
    Next i
End Sub

' --- 検索ボタン ---
Private Sub SearchButton_Click()
    LoadParticipantList Trim(Me.SearchBox.Value)
    If Me.ListBox1.ListCount = 0 Then
        MsgBox "該当者が見つかりませんでした。", vbInformation
    End If
End Sub

' --- クリアボタン ---
Private Sub ClearButton_Click()
    Me.SearchBox.Value = ""
    Me.NavLabel.Caption = ""
    LoadParticipantList
End Sub

' --- 選択ボタン：伝票一覧の一致セルを収集して最初へジャンプ ---
Private Sub SelectButton_Click()
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "対象を選択してください。", vbExclamation
        Exit Sub
    End If

    Dim searchValue As String
    Dim searchCol As Long
    
    If Me.OptionBuyer.Value Then
        searchCol = COL_BUYER
        searchValue = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
    Else
        searchCol = COL_SELLER
        searchValue = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
    End If

    CollectMatches searchValue, searchCol

    If matchCount = 0 Then
        MsgBox "伝票一覧に一致するデータが見つかりませんでした。", vbInformation
        Me.NavLabel.Caption = ""
        Exit Sub
    End If

    currentIndex = 0
    GoToCurrentCell
    UpdateNavLabel
End Sub

' --- 伝票一覧シートから一致セルのアドレスを全収集 ---
Private Sub CollectMatches(ByVal number As String, ByVal col As Long)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, count As Long

    Set ws = Worksheets("伝票一覧")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    count = 0

    ReDim matchCells(1 To lastRow)
    
    For i = 2 To lastRow
        If CStr(ws.Cells(i, col).Value) = number Then
            count = count + 1
            matchCells(count) = ws.Cells(i, col).Address
        End If
    Next i
    
    matchCount = count
    If matchCount > 0 Then ReDim Preserve matchCells(1 To matchCount)
End Sub

' --- 現在インデックスのセルへジャンプ ---
Private Sub GoToCurrentCell()
    With Worksheets("伝票一覧")
        .Activate
        .Range(matchCells(currentIndex + 1)).Select
    End With
End Sub

' --- 次へ ---
Private Sub NextButton_Click()
    If matchCount = 0 Then Exit Sub
    currentIndex = (currentIndex + 1) Mod matchCount
    GoToCurrentCell
    UpdateNavLabel
End Sub

' --- 前へ ---
Private Sub PrevButton_Click()
    If matchCount = 0 Then Exit Sub
    currentIndex = (currentIndex - 1 + matchCount) Mod matchCount
    GoToCurrentCell
    UpdateNavLabel
End Sub

' --- 件数ラベル更新 ---
Private Sub UpdateNavLabel()
    Me.NavLabel.Caption = (currentIndex + 1) & " / " & matchCount & " 件"
End Sub

' --- 閉じるボタン ---
Private Sub CloseButton_Click()
    Unload Me
End Sub
