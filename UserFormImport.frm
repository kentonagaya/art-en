Option Explicit

' --- 売主検索ボタン ---
Private Sub SearchButtonSeller_Click()

    ' 売主検索フォームを表示
    With UserFormSearch
        Set .ParentForm = Me
        .SearchMode = "SELLER"
        .Show vbModal
    End With

End Sub

' --- 買主検索ボタン ---
Private Sub SearchButtonBuyer_Click()

    ' 買主検索フォームを表示
    With UserFormSearch
        Set .ParentForm = Me
        .SearchMode = "BUYER"
        .Show vbModal
    End With

End Sub


' --- 次番号設定 ---
Private Sub SetNextSlipNumberButton_Click()

    SetNextSlipNumber

End Sub


' --- リスト初期化 ---
Private Sub UserForm_Initialize()

    SetNextSlipNumber

End Sub

' --- 伝票番号設定共通処理 ---
Private Sub SetNextSlipNumber()

    Dim lastRow    As Long
    Dim lastVal    As String
    Dim prefix     As String
    Dim seqStr     As String
    Dim seqNum     As Long
    Dim seqLen     As Long
    Dim i          As Long

    With Worksheets("伝票一覧")

        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row

        ' データなし（見出し行のみ）の場合
        If .Cells(lastRow, 1).Value = "伝票番号" Then
            Me.SlipNumber.Text = "1"
            Exit Sub
        End If

        lastVal = CStr(.Cells(lastRow, 1).Value)

        ' --- 形式判定：数値のみ か 英数字混在か ---
        If IsNumeric(lastVal) Then

            ' 旧形式：数値のみ → +1
            Me.SlipNumber.Text = CLng(lastVal) + 1

        Else

            ' 新形式：英数字混在 → 末尾の連番部分のみインクリメント
            ' 末尾から数字を取り出す
            seqStr = ""
            For i = Len(lastVal) To 1 Step -1
                If Mid(lastVal, i, 1) Like "[0-9]" Then
                    seqStr = Mid(lastVal, i, 1) & seqStr
                Else
                    Exit For
                End If
            Next i

            ' プレフィックス（日付＋担当者コード）を取り出す
            prefix = Left(lastVal, Len(lastVal) - Len(seqStr))
            seqLen = Len(seqStr)
            seqNum = CLng(seqStr) + 1

            ' 連番をゼロ埋めして再結合
            Me.SlipNumber.Text = prefix & Format(seqNum, String(seqLen, "0"))

        End If

    End With

End Sub

' --- 登録済みのうち最新の番号の取得 ---
Private Sub GetLatestInformationButton_Click()

    Dim lastRow As Long
    Dim rawItem As String
    Dim ItemName As String
    Dim itemCount As String
    Dim i As Long
    Dim c As String

    With Worksheets("伝票一覧")

        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row

        Me.SlipNumber.Text = .Cells(lastRow, 1).Value

        rawItem = .Cells(lastRow, 2).Value
        itemCount = ""

        For i = Len(rawItem) To 1 Step -1
            c = Mid(rawItem, i, 1)
            If c Like "[0-9]" Then
                itemCount = c & itemCount
            ElseIf c = "点" Then
            Else
                Exit For
            End If
        Next i

        Me.ItemName.Text = Trim(Left(rawItem, i))
        Me.Quantity.Text = itemCount
        Me.Price.Text = CLng(.Cells(lastRow, 3).Value) \ 1000
        Me.BuyerNumber.Text = .Cells(lastRow, 5).Value
        Me.BuyerName.Text = .Cells(lastRow, 4).Value
        Me.SellerNumber.Text = .Cells(lastRow, 7).Value
        Me.SellerName.Text = .Cells(lastRow, 6).Value

    End With

End Sub

'--- リストマスタの内容を「項目」に反映 ---
Private Sub RefreshListButton_Click()

    Const MASTER_FILE_NAME As String = "御堂会顧客リストマスタ.xlsx"
    Dim srcWorkbook As Workbook
    Dim destWorkbook As Workbook
    Dim srcWorksheet As Worksheet
    Dim destWorksheet As Worksheet
    Dim filePath As String
    Dim sep As String: sep = Application.PathSeparator ' OSごとの区切り文字
     
    ' 開いているブック（更新先）
    Set destWorkbook = ThisWorkbook
    Set destWorksheet = destWorkbook.Sheets("顧客一覧")
    
    ' 顧客リストマスタのパス
    filePath = ThisWorkbook.Path & sep & MASTER_FILE_NAME
    
    ' ファイル存在チェック
    If Dir(filePath) = "" Then
        MsgBox "顧客リストマスタが見つかりません。" & vbCrLf & _
               "以下の場所を確認してください。" & vbCrLf & _
               filePath, vbExclamation
        Exit Sub
    End If
    
    ' 顧客リストマスタブックを開く
    Set srcWorkbook = Workbooks.Open(filePath)
    Set srcWorksheet = srcWorkbook.Sheets(1) ' 必要に応じて変更
    
    ' コピー
    srcWorksheet.Cells.Copy destWorksheet.Cells
    
    ' 保存せずに閉じる
    srcWorkbook.Close SaveChanges:=False
    
    MsgBox "顧客リストを更新しました。", vbInformation

End Sub

'--- 品名に「掛軸」を記入 ---
Private Sub KakejikuButton_Click()

    Me.ItemName.Text = Me.ItemName.Text + "掛軸 "

End Sub

'--- 品名に「マクリ」を記入 ---
Private Sub MakuriButton_Click()

    Me.ItemName.Text = Me.ItemName.Text + "マクリ "

End Sub

'--- 品名に「巻物」を記入 ---
Private Sub MakimonoButton_Click()

    Me.ItemName.Text = Me.ItemName.Text + "巻物 "

End Sub

'--- 品名に「画帖」を記入 ---
Private Sub GajoButton_Click()

    Me.ItemName.Text = Me.ItemName.Text + "画帖 "

End Sub

'--- 品名に「山」を記入 ---
Private Sub YamaButton_Click()

    Me.ItemName.Text = Me.ItemName.Text + "山 "

End Sub

'--- 入力項目全削除ボタン ---
Private Sub AllClearButton_Click()
    
    ClearInputs False

End Sub

'--- 入力項目を「マスタ」に登録 ---
Private Sub RegisterButton_Click()
 
    If Not ValidateAndResolveParties Then Exit Sub
    RegisterToMaster True
    ShowRegisterMessage
    ClearInputs False
    ActiveWorkbook.Save

End Sub

'--- 入力項目を「マスタ」に登録（売主情報継続） ---
Private Sub KeepInfoRegisterButton_Click()

    If Not ValidateAndResolveParties Then Exit Sub
    RegisterToMaster False
    ShowRegisterMessage
    ClearInputs True
    ActiveWorkbook.Save
  
End Sub

' --- 入力内容確認共通処理 ---
Private Function ValidateAndResolveParties() As Boolean

    Dim ws As Worksheet
    Dim myRange As Range, myObj As Range
    Dim keyword As String

    Set ws = Worksheets("顧客一覧")
    ValidateAndResolveParties = False

    ' 金額
    If Me.Price.Value = "" Then
        MsgBox "金額未入力"
        Exit Function
    End If

    ' --- 買主 ---
    If Me.BuyerNumber.Value = "" Then
        If Me.BuyerName.Value = "" Then
            MsgBox "買い未入力です"
            Exit Function
        End If

        Set myRange = ws.Range("B2:B500")
        keyword = Me.BuyerName.Value
        Set myObj = myRange.Find(keyword, LookAt:=xlWhole)

        If myObj Is Nothing Then
            MsgBox "買主未入力または '" & keyword & "' はありません"
            Exit Function
        End If

        Me.BuyerNumber.Value = ws.Cells(myObj.Row, 1).Value

    Else
        Set myRange = ws.Range("A2:A500")
        keyword = Me.BuyerNumber.Value
        Set myObj = myRange.Find(keyword, LookAt:=xlWhole)

        If myObj Is Nothing Then
            MsgBox "買主番号 " & keyword & " はありません"
            Exit Function
        End If

        Me.BuyerName.Value = ws.Cells(myObj.Row, 2).Value
    End If

    ' --- 売主 ---
    If Me.SellerNumber.Value = "" Then
        If Me.SellerName.Value = "" Then
            MsgBox "売り未入力です"
            Exit Function
        End If

        Set myRange = ws.Range("C2:C500")
        keyword = Me.SellerName.Value
        Set myObj = myRange.Find(keyword, LookAt:=xlWhole)

        If myObj Is Nothing Then
            MsgBox "売主未入力または '" & keyword & "' はありません"
            Exit Function
        End If

        Me.SellerNumber.Value = ws.Cells(myObj.Row, 1).Value

    Else
        Set myRange = ws.Range("A2:A500")
        keyword = Me.SellerNumber.Value
        Set myObj = myRange.Find(keyword, LookAt:=xlWhole)

        If myObj Is Nothing Then
            MsgBox "売主番号 " & keyword & " はありません"
            Exit Function
        End If

        Me.SellerName.Value = ws.Cells(myObj.Row, 3).Value
    End If

    ' その他必須
    If Me.ItemName.Value = "" Or Me.SlipNumber.Value = "" Then
        MsgBox "品名または番号未入力"
        Exit Function
    End If

    '--- 相互売買警告 ---
    If IsMutualTradeWarning(Me.BuyerNumber.Text, Me.SellerNumber.Text) Then
        MsgBox _
            "【相互売買警告】" & vbCrLf & _
            "この売主・買主の組み合わせは登録できません。" & vbCrLf & _
            "売主：" & Me.SellerNumber.Text & " " & Me.SellerName.Text & vbCrLf & _
            "買主：" & Me.BuyerNumber.Text & " " & Me.BuyerName.Text, _
            vbCritical
        Exit Function
    End If

    ValidateAndResolveParties = True

End Function

' --- マスタ登録共通処理 ---
Private Sub RegisterToMaster(ByVal AllowOverwrite As Boolean)

    Dim ws As Worksheet
    Dim lastRow As Long, nn As Long

    Set ws = Worksheets("伝票一覧")

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

    If AllowOverwrite And ws.Cells(lastRow, 1).Value = Me.SlipNumber.Text Then
        If MsgBox("上書きですか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        nn = lastRow
    Else
        nn = lastRow + 1
    End If

    ws.Cells(nn, 1).Value = Me.SlipNumber.Text
    ws.Cells(nn, 2).Value = Me.ItemName.Text & Me.Quantity.Text & "点"
    ws.Cells(nn, 3).Value = CLng(Me.Price.Value) * 1000
    ws.Cells(nn, 4).Value = Me.BuyerName.Text
    ws.Cells(nn, 5).Value = Me.BuyerNumber.Text
    ws.Cells(nn, 6).Value = Me.SellerName.Text
    ws.Cells(nn, 7).Value = Me.SellerNumber.Text
    ws.Cells(nn, 8).Value = Time

End Sub

' --- 登録完了メッセージ表示共通処理 ---
Private Sub ShowRegisterMessage()

    MsgBox _
        "登録しました。" & vbCrLf & vbCrLf & _
        "伝票番号：" & Me.SlipNumber.Text & vbCrLf & _
        "品名　　：" & Me.ItemName.Text & Me.Quantity.Text & "点" & vbCrLf & _
        "落札額　：" & Format((CLng(Me.Price.Value) * 1000), "#,##0") & "円" & vbCrLf & _
        "売主　　：" & Me.SellerNumber.Text & " " & Me.SellerName.Text & vbCrLf & _
        "買主　　：" & Me.BuyerNumber.Text & " " & Me.BuyerName.Text, _
        vbInformation, "登録完了"

End Sub


' --- 入力内容全削除共通処理（売主情報継続判断） ---
Private Sub ClearInputs(Optional ByVal KeepSeller As Boolean = False)

    SetNextSlipNumber
    Me.BuyerNumber.Text = ""
    Me.BuyerName.Text = ""
    Me.ItemName.Text = ""
    Me.Price.Text = ""
    Me.Quantity.Text = ""

    If Not KeepSeller Then
        Me.SellerNumber.Text = ""
        Me.SellerName.Text = ""
    End If

End Sub

' --- 閉じるボタン ---
Private Sub CloseButton_Click()

    Unload UserFormImport

End Sub
