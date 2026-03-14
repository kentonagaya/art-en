Option Explicit

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

' --- リスト初期化 ---
Private Sub UserForm_Initialize()

    LoadParticipantList

End Sub

' --- 参加者リスト読み込み・検索対応 ---
Private Sub LoadParticipantList(Optional ByVal keyword As String = "")

    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim clientNameStr As String, shortNameStr As String, clientNumberStr As String

    Set ws = Worksheets("参加者一覧")

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

            Me.ListBox1.AddItem ws.Cells(i, 1).Value ' 1列目：参加者番号
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = ws.Cells(i, 2).Value ' 2列目：屋号
            Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = ws.Cells(i, 3).Value ' 3列目：売りカタ

        End If

    Next i

End Sub

' --- 閉じるボタン ---
Private Sub CloseButton_Click()

    Unload UserFormExport
    
End Sub

' --- 仮伝票ボタン ---
Private Sub ProvisionalExportButton_Click()

    CreateVoucher IsProvisional:=True
    
End Sub

' --- 伝票出力ボタン ---
Private Sub ExportButton_Click()

    CreateVoucher IsProvisional:=False
    
End Sub

' --- 出力共通処理 ---
Private Sub CreateVoucher(ByVal IsProvisional As Boolean)
    
    Const START_ROW As Long = 17
    Const FIRST_PAGE_ROWS As Long = 33
    Const NEXT_PAGE_ROWS As Long = 52
    Const BORDER_START_COL As Long = 1  ' A列
    Const BORDER_END_COL As Long = 8 ' H列
    Dim clientName As String, shortName As String, clientNumber As String
    Dim desktopPath As String, dateStr As String
    Dim lRow As Long, mRow As Long, i As Long, pageCount As Long
    Dim lastSellRow As Long, lastBuyRow As Long, lastDataRow As Long, remainRows As Long
    Dim pageIndex As Long, pageStartRow As Long, pageEndRow As Long
    Dim ws01 As Worksheet, ws02 As Worksheet, ws03 As Worksheet
    Dim newBook As Workbook
    Dim exportSheets As Variant
    Dim sep As String: sep = Application.PathSeparator ' OSごとの区切り文字
    
    ' 入力チェック
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "出力する対象を選択してください。", vbExclamation
        Exit Sub
    End If

    ' データの取得
    clientNumber = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
    clientName = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
    shortName = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
    
    ' シートのコピーと名前変更
    Sheets("伝票テンプレート").Copy After:=Sheets(Sheets.Count)
    Set ws02 = ActiveSheet
    
    ' 同名のシートがある場合のエラー対策
    On Error Resume Next
    ws02.Name = clientName
    If Err.Number <> 0 Then
        MsgBox "既に同名のシートが存在します。処理を中断します。"
        Application.DisplayAlerts = False
        ws02.Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If
    On Error GoTo 0
    
    ' データの転記
    Set ws01 = Worksheets("伝票一覧")
    ws02.Cells(1, 1).Value = clientNumber
    ws02.Cells(1, 2).Value = clientName
    
    lRow = ws01.Cells(ws01.Rows.Count, "A").End(xlUp).Row
    
    ' --- 売り側の転記 ---
    mRow = START_ROW
    For i = 2 To lRow
        If ws01.Cells(i, 6).Value = shortName Then
            ws02.Cells(mRow, 1).Resize(1, 4).Value = _
                Array(ws01.Cells(i, 1), ws01.Cells(i, 2), ws01.Cells(i, 3), ws01.Cells(i, 4))
            
            If ws01.Cells(i, 4).Value = "引き" Then
                ws02.Cells(mRow, 1).Resize(1, 4).Interior.ColorIndex = 3
                MsgBox "引きの荷物お忘れなく!!"
            End If
            mRow = mRow + 1
        End If
    Next i
    ws02.Cells(mRow, 2).Value = "以下余白"
       
    ' --- 買い側の転記 ---
    mRow = START_ROW
    For i = 2 To lRow
        If ws01.Cells(i, 4).Value = clientName Then
            ws02.Cells(mRow, 5).Resize(1, 4).Value = _
                Array(ws01.Cells(i, 1), ws01.Cells(i, 2), ws01.Cells(i, 3), ws01.Cells(i, 6))
            mRow = mRow + 1
        End If
    Next i
    ws02.Cells(mRow, 6).Value = "以下余白"
    
    ' --- 仮伝票の場合の追加処理 ---
    If IsProvisional Then
        ws02.Cells(1, 7).Value = "仮伝票"
        ws02.Cells(3, 1).Value = "仮伝票です。精算できません。"
    End If


    ' --- 保存と書き出し処理 ---
    ' ページ数を行数ベースで算出
    lastSellRow = ws02.Cells(ws02.Rows.Count, 1).End(xlUp).Row
    lastBuyRow = ws02.Cells(ws02.Rows.Count, 5).End(xlUp).Row
    
    lastDataRow = Application.WorksheetFunction.Max(lastSellRow, lastBuyRow)
    
    If lastDataRow <= START_ROW + FIRST_PAGE_ROWS - 1 Then
        pageCount = 1
    Else
        remainRows = lastDataRow - (START_ROW + FIRST_PAGE_ROWS - 1)
        pageCount = 1 + Application.WorksheetFunction.RoundUp( _
                        remainRows / NEXT_PAGE_ROWS, 0)
    End If
    
    ws02.PageSetup.RightHeader = "全" & pageCount & "枚"
    
    ' 罫線の指定
    For pageIndex = 1 To pageCount
        
        If pageIndex = 1 Then
            ' --- 1枚目 ---
            pageStartRow = START_ROW
            pageEndRow = START_ROW + FIRST_PAGE_ROWS - 1
        Else
            ' --- 2枚目以降 ---
            pageStartRow = START_ROW + FIRST_PAGE_ROWS _
                           + (pageIndex - 2) * NEXT_PAGE_ROWS
            pageEndRow = pageStartRow + NEXT_PAGE_ROWS - 1
        End If
        
        With ws02.Range( _
            ws02.Cells(pageStartRow, BORDER_START_COL), _
            ws02.Cells(pageEndRow, BORDER_END_COL) _
        ).Borders
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    Next pageIndex

    ' 伝票出力用のみ、控えと荷渡し用を作成
    If Not IsProvisional Then
        ' --- 1. 控え作成 ---
        ws02.Copy After:=ws02
        Set ws03 = ws02.Parent.Sheets(ws02.Index + 1)
        ws03.Name = "控え"
        ws03.Cells(1, 7).Value = "事務局控"


        ' --- 2. 荷渡し用作成 ---
        ws02.Copy After:=ws03
        Set ws03 = ws02.Parent.Sheets(ws02.Index + 2)
        ws03.Name = "荷渡し用"
        ws03.Cells(1, 7).Value = "荷蔵用"

        exportSheets = Array(ws02.Name, "控え", "荷渡し用")
    
    Else
    
        ' 仮伝票は1枚のみ
        exportSheets = Array(ws02.Name)
    
    End If
    
    ws02.Parent.Sheets(exportSheets).Move
    Set newBook = ActiveWorkbook
    
    ' OS対応の保存処理・警告の一時非表示
    Application.DisplayAlerts = False
    dateStr = Format(Date, "yyyymmdd")
    newBook.SaveAs _
        Filename:=ThisWorkbook.Path & sep & clientName & "_" & dateStr & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    Unload UserFormExport
    MsgBox "No." & clientNumber & " " & clientName & " の伝票ファイルが作成されました。"
    
End Sub



