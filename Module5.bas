Option Explicit

Sub 集計1()

    Const START_ROW As Long = 5

    Dim wsMaster As Worksheet
    Dim wsItem   As Worksheet
    Dim wsResult As Worksheet
    Dim dataArr  As Variant
    Dim itemArr  As Variant
    Dim i As Long, j As Long
    Dim outRow As Long

    Set wsMaster = Worksheets("伝票一覧")
    Set wsItem   = Worksheets("参加者一覧")
    Set wsResult = Worksheets("集計1")

    ' --- 集計1シートをクリア ---
    wsResult.Range("A" & START_ROW & ":H999").Clear

    ' --- データ一括取得 ---
    If wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row < 2 Or _
       wsItem.Cells(wsItem.Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "データがありません。", vbExclamation
        Exit Sub
    End If

    dataArr = wsMaster.Range("A1").CurrentRegion.Value
    itemArr = wsItem.Range("A1").CurrentRegion.Value

    outRow = START_ROW

    For i = 2 To UBound(itemArr, 1)

        Dim pName  As String  ' 屋号（参加者一覧B列）
        Dim pShort As String  ' 売りカタ（参加者一覧C列）
        Dim pRate  As Double  ' 手数料率

        pName  = CStr(itemArr(i, 2))
        pShort = CStr(itemArr(i, 3))

        ' --- 手数料率の決定 ---
        ' E列が空欄→5%、数値あり（0含む）→その値を使用
        If IsEmpty(itemArr(i, 5)) Or itemArr(i, 5) = "" Then
            pRate = 0.05
        ElseIf IsNumeric(itemArr(i, 5)) Then
            pRate = CDbl(itemArr(i, 5))
            If pRate > 1 Then pRate = pRate / 100  ' 「5」→0.05、「10」→0.10 に変換
        Else
            pRate = 0.05
        End If

        ' --- 買い側集計（伝票一覧D列 = 屋号で照合） ---
        Dim buyTotal As Double
        buyTotal = 0
        For j = 2 To UBound(dataArr, 1)
            If CStr(dataArr(j, 4)) = pName Then
                buyTotal = buyTotal + CDbl(dataArr(j, 3))
            End If
        Next j

        ' --- 売り側集計（伝票一覧F列 = 売りカタで照合） ---
        Dim sellTotal As Double
        sellTotal = 0
        For j = 2 To UBound(dataArr, 1)
            If CStr(dataArr(j, 6)) = pShort Then
                sellTotal = sellTotal + CDbl(dataArr(j, 3))
            End If
        Next j

        ' --- 取引なしはスキップ ---
        If buyTotal = 0 And sellTotal = 0 Then GoTo NextParticipant

        ' --- 出力 ---
        With wsResult
            .Cells(outRow, 1).Value = pName                          ' A: 買主（屋号）
            .Cells(outRow, 2).Value = buyTotal * (1 + pRate)         ' B: 買い側合計（落札額＋手数料）
            .Cells(outRow, 3).Value = buyTotal                       ' C: 買い落札額
            .Cells(outRow, 4).Value = buyTotal * pRate               ' D: 買い手数料

            .Cells(outRow, 5).Value = pShort                         ' E: 売りカタ
            .Cells(outRow, 6).Value = sellTotal * (1 - pRate)        ' F: 売り側合計（落札額－手数料）
            .Cells(outRow, 7).Value = sellTotal                      ' G: 売り落札額
            .Cells(outRow, 8).Value = sellTotal * pRate              ' H: 売り手数料
        End With

        outRow = outRow + 1

NextParticipant:
    Next i

    MsgBox "集計が完了しました。", vbInformation

End Sub