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

    Dim lastRow As Long

    With Worksheets("マスタ")

        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

        If .Cells(lastRow, 1).Value = "伝票番号" Then
            Me.番号.Text = "1"
        Else
            Me.番号.Text = .Cells(lastRow, 1).Value + 1
        End If

    End With

End Sub

' --- 登録済みのうち最新の番号の取得 ---
Private Sub GetLatestInformationButton_Click()

    Dim lastRow As Long
    
    With Worksheets("マスタ")
        
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
         
        Me.番号.Text = .Cells(lastRow, 1).Value
        Me.品名.Text = .Cells(lastRow, 2).Value
        Me.金額.Text = .Cells(lastRow, 3).Value
          
        Me.買主.Text = .Cells(lastRow, 5).Value
        Me.かいな.Text = .Cells(lastRow, 4).Value
        Me.売主.Text = .Cells(lastRow, 7).Value
        Me.うりな.Text = .Cells(lastRow, 6).Value
        
    End With

End Sub


'--- リストマスタの内容を「項目」に反映 ---
Private Sub 顧客更新_Click()

    Const MASTER_FILE_NAME As String = "御堂会顧客リストマスタ.xlsx"
    Dim srcWorkbook As Workbook
    Dim destWorkbook As Workbook
    Dim srcWorksheet As Worksheet
    Dim destWorksheet As Worksheet
    Dim filePath As String
    Dim sep As String: sep = Application.PathSeparator ' OSごとの区切り文字
     
    ' 開いているブック（更新先）
    Set destWorkbook = ThisWorkbook
    Set destWorksheet = destWorkbook.Sheets("項目")
    
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
Private Sub 掛軸_Click()

    Me.品名.Text = Me.品名.Text + "掛軸 "

End Sub

'--- 品名に「マクリ」を記入 ---
Private Sub マクリ_Click()

    Me.品名.Text = Me.品名.Text + "マクリ"

End Sub

'--- 品名に「巻物」を記入 ---
Private Sub 巻物_Click()

    Me.品名.Text = Me.品名.Text + "巻物"

End Sub

'--- 品名に「画帖」を記入 ---
Private Sub 画帖_Click()

    Me.品名.Text = Me.品名.Text + "画帖"

End Sub

'--- 品名の値を削除 ---
Private Sub ClearItemButton_Click()

    Me.品名.Value = ""

End Sub

'--- 個数に「点」を記入 ---
Private Sub 点_Click()

    Me.個数.Text = Me.個数.Text + "点"

End Sub

'--- 個数に「山」を記入 ---
Private Sub 山_Click()

    Me.個数.Text = Me.個数.Text + "山"

End Sub

Private Sub 金額_Change()

    'If IsNumeric(Me.金額.Text) Then
    '    Me.金額.Text = Format(Me.金額.Text, "#,##0")
    'End If

End Sub

'--- 金額の値を削除 ---
Private Sub 値段クリア_Click()

    Me.金額.Value = ""

End Sub

'--- 入力項目全削除ボタン ---
Private Sub AllClearButton_Click()
    
    ClearInputs False

End Sub

'--- 入力項目を「マスタ」に登録 ---
Private Sub 登録_Click()
 
    If Not ValidateAndResolveParties Then Exit Sub
    RegisterToMaster True
    ShowRegisterMessage
    ClearInputs False
    ActiveWorkbook.Save

End Sub

'--- 入力項目を「マスタ」に登録（売主情報継続） ---
Private Sub 売主継続登録_Click()

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

    Set ws = Worksheets("項目")
    ValidateAndResolveParties = False

    ' 金額
    If Me.金額.Value = "" Then
        MsgBox "金額未入力"
        Exit Function
    End If

    ' --- 買主 ---
    If Me.買主.Value = "" Then
        If Me.かいな.Value = "" Then
            MsgBox "買い未入力です"
            Exit Function
        End If

        Set myRange = ws.Range("B2:B500")
        keyword = Me.かいな.Value
        Set myObj = myRange.Find(keyword, LookAt:=xlWhole)

        If myObj Is Nothing Then
            MsgBox "買主未入力または '" & keyword & "' はありません"
            Exit Function
        End If

        Me.買主.Value = ws.Cells(myObj.Row, 1).Value

    Else
        Set myRange = ws.Range("A2:A500")
        keyword = Me.買主.Value
        Set myObj = myRange.Find(keyword, LookAt:=xlWhole)

        If myObj Is Nothing Then
            MsgBox "買主番号 " & keyword & " はありません"
            Exit Function
        End If

        Me.かいな.Value = ws.Cells(myObj.Row, 2).Value
    End If

    ' --- 売主 ---
    If Me.売主.Value = "" Then
        If Me.うりな.Value = "" Then
            MsgBox "売り未入力です"
            Exit Function
        End If

        Set myRange = ws.Range("C2:C500")
        keyword = Me.うりな.Value
        Set myObj = myRange.Find(keyword, LookAt:=xlWhole)

        If myObj Is Nothing Then
            MsgBox "売主未入力または '" & keyword & "' はありません"
            Exit Function
        End If

        Me.売主.Value = ws.Cells(myObj.Row, 1).Value

    Else
        Set myRange = ws.Range("A2:A500")
        keyword = Me.売主.Value
        Set myObj = myRange.Find(keyword, LookAt:=xlWhole)

        If myObj Is Nothing Then
            MsgBox "売主番号 " & keyword & " はありません"
            Exit Function
        End If

        Me.うりな.Value = ws.Cells(myObj.Row, 3).Value
    End If

    ' その他必須
    If Me.品名.Value = "" Or Me.番号.Value = "" Then
        MsgBox "品名または番号未入力"
        Exit Function
    End If

    '--- 相互売買警告 ---
    If IsMutualTradeWarning(Me.買主.Text, Me.売主.Text) Then
        MsgBox _
            "【相互売買警告】" & vbCrLf & _
            "この売主・買主の組み合わせは登録できません。" & vbCrLf & _
            "売主：" & Me.売主.Text & " " & Me.うりな.Text & vbCrLf & _
            "買主：" & Me.買主.Text & " " & Me.かいな.Text, _
            vbCritical
        Exit Function
    End If

    ValidateAndResolveParties = True

End Function

' --- マスタ登録共通処理 ---
Private Sub RegisterToMaster(ByVal AllowOverwrite As Boolean)

    Dim ws As Worksheet
    Dim lastRow As Long, nn As Long

    Set ws = Worksheets("マスタ")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If AllowOverwrite And ws.Cells(lastRow, 1).Value = Me.番号.Text Then
        If MsgBox("上書きですか？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        nn = lastRow
    Else
        nn = lastRow + 1
    End If

    ws.Cells(nn, 1).Value = Me.番号.Text
    ws.Cells(nn, 2).Value = Me.品名.Text & Me.個数.Text
    ws.Cells(nn, 3).Value = CLng(Me.金額.Value) * 1000
    ws.Cells(nn, 4).Value = Me.かいな.Text
    ws.Cells(nn, 5).Value = Me.買主.Text
    ws.Cells(nn, 6).Value = Me.うりな.Text
    ws.Cells(nn, 7).Value = Me.売主.Text
    ws.Cells(nn, 9).Value = Time

End Sub

' --- 登録完了メッセージ表示共通処理 ---
Private Sub ShowRegisterMessage()

    MsgBox _
        "登録しました。" & vbCrLf & vbCrLf & _
        "伝票番号：" & Me.番号.Text & vbCrLf & _
        "品名　　：" & Me.品名.Text & Me.個数.Text & vbCrLf & _
        "落札額　：" & Format((CLng(Me.金額.Value) * 1000), "#,##0") & "円" & vbCrLf & _
        "売主　　：" & Me.売主.Text & " " & Me.うりな.Text & vbCrLf & _
        "買主　　：" & Me.買主.Text & " " & Me.かいな.Text, _
        vbInformation, "登録完了"

End Sub


' --- 入力内容全削除共通処理（売主情報継続判断） ---
Private Sub ClearInputs(Optional ByVal KeepSeller As Boolean = False)

    SetNextSlipNumber
    Me.買主.Text = ""
    Me.かいな.Text = ""
    Me.品名.Text = ""
    Me.金額.Text = ""
    Me.個数.Text = ""

    If Not KeepSeller Then
        Me.売主.Text = ""
        Me.うりな.Text = ""
    End If

End Sub

' --- 閉じるボタン ---
Private Sub CloseButton_Click()

    Unload UserFormImport

End Sub


