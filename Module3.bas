Option Explicit

' --- 名前自動入力機能（マスタシート） ---
Public Sub 名前自動入力()

    Dim wsMaster As Worksheet
    Dim wsItem As Worksheet
    Dim targetCell As Range
    Dim findRange As Range
    Dim foundCell As Range
    
    Dim clientNumber As String
    Dim clientName As String
    Dim isBuyer As Boolean
    Dim isSeller As Boolean

    Set wsMaster = Worksheets("伝票一覧")
    Set wsItem = Worksheets("参加者一覧")

    '--- Selection がセルか確認 ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "セルを選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    Set targetCell = Selection.Cells(1)

    '--- マスタシート上か確認 ---
    If targetCell.Parent.Name <> "伝票一覧" Then
        MsgBox "マスタシート上のセルを選択してください。", vbExclamation
        Exit Sub
    End If

    '--- 列判定 ---
    Select Case targetCell.Column
        Case 5   ' E列：買主番号
            isBuyer = True
        Case 7   ' G列：売主番号
            isSeller = True
        Case Else
            MsgBox "番号セルを選択してください。", vbExclamation
            Exit Sub
    End Select

    '--- 番号チェック ---
    If Trim(targetCell.Value) = "" Then
        MsgBox "番号が入力されていません。", vbExclamation
        Exit Sub
    End If

    clientNumber = CStr(targetCell.Value)

    '--- 項目シート検索 ---
    Set findRange = wsItem.Range("A:A")
    Set foundCell = findRange.Find(What:=clientNumber, LookAt:=xlWhole)

    If foundCell Is Nothing Then
        MsgBox "番号「" & clientNumber & "」は項目シートに存在しません。", vbCritical
        Exit Sub
    End If

    '--- 名前取得 ---
    If isBuyer Then
        clientName = wsItem.Cells(foundCell.Row, 2).Value   ' B列：買主名前
        wsMaster.Cells(targetCell.Row, 4).Value = clientName   ' D列
    ElseIf isSeller Then
        clientName = wsItem.Cells(foundCell.Row, 3).Value   ' C列：売主名前
        wsMaster.Cells(targetCell.Row, 6).Value = clientName   ' F列
    End If

End Sub

' --- 名前自動入力機能（相互売買警告シート） ---
Public Sub 名前自動入力B()

    Dim wsWarn As Worksheet
    Dim wsItem As Worksheet
    Dim targetCell As Range
    Dim findRange As Range
    Dim foundCell As Range
    
    Dim clientNumber As String
    Dim clientName As String
    Dim shortName As String
    
    Dim isLeft As Boolean
    Dim isRight As Boolean

    Set wsWarn = Worksheets("相互売買警告リスト")
    Set wsItem = Worksheets("参加者一覧")

    '--- Selection がセルか確認 ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "セルを選択してから実行してください。", vbExclamation
        Exit Sub
    End If

    Set targetCell = Selection.Cells(1)

    '--- シートと列チェック ---
    If targetCell.Parent.Name <> "相互売買警告リスト" Then
        MsgBox "正しいシートを選択してください。", vbExclamation
        Exit Sub
    End If
    
    '--- 列判定 ---
    Select Case targetCell.Column
        Case 1   ' A列：参加者A
            isLeft = True
        Case 4   ' G列：参加者B
            isRight = True
        Case Else
            MsgBox "番号セルを選択してください。", vbExclamation
            Exit Sub
    End Select

    '--- 番号チェック ---
    If Trim(targetCell.Value) = "" Then
        MsgBox "番号が入力されていません。", vbExclamation
        Exit Sub
    End If

    clientNumber = CStr(targetCell.Value)

    '--- 項目シート検索 ---
    Set findRange = wsItem.Range("A:A")
    Set foundCell = findRange.Find(What:=clientNumber, LookAt:=xlWhole)

    If foundCell Is Nothing Then
        MsgBox "番号「" & clientNumber & "」は項目シートに存在しません。", vbCritical
        Exit Sub
    End If

    '--- 名前取得 ---
    clientName = wsItem.Cells(foundCell.Row, 2).Value ' 例：B列（屋号）
    shortName = wsItem.Cells(foundCell.Row, 3).Value ' 例：C列（売りカタ）
        
    If isLeft Then
        wsWarn.Cells(targetCell.Row, 2).Value = clientName
        wsWarn.Cells(targetCell.Row, 3).Value = shortName
    ElseIf isRight Then
        wsWarn.Cells(targetCell.Row, 5).Value = clientName
        wsWarn.Cells(targetCell.Row, 6).Value = shortName
    End If

End Sub


