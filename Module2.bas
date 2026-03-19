Option Explicit

Sub 集計()

    Const START_ROW As Long = 10

    Dim wsMaster As Worksheet
    Dim wsItem   As Worksheet
    Dim wsResult As Worksheet
    Dim dataArr  As Variant
    Dim itemArr  As Variant
    Dim i As Long, j As Long
    Dim outRow As Long

    Set wsMaster = Worksheets("伝票一覧")
    Set wsItem = Worksheets("参加者一覧")
    Set wsResult = Worksheets("集計")

    ' --- クリア ---
    wsResult.Range("A" & START_ROW & ":H999").Clear
    wsResult.Range("B2:B4").Clear
    wsResult.Range("B6:B7").Clear

    If wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row < 2 Or _
       wsItem.Cells(wsItem.Rows.count, 1).End(xlUp).Row < 2 Then
        MsgBox "データがありません。", vbExclamation
        Exit Sub
    End If

    dataArr = wsMaster.Range("A1").CurrentRegion.Value
    itemArr = wsItem.Range("A1").CurrentRegion.Value

    outRow = START_ROW

    For i = 2 To UBound(itemArr, 1)

        Dim pNum   As String
        Dim pName  As String
        Dim pShort As String
        Dim pRate  As Double

        pNum = CStr(itemArr(i, 1))
        pName = CStr(itemArr(i, 2))
        pShort = CStr(itemArr(i, 3))

        ' --- 手数料率の決定 ---
        If IsEmpty(itemArr(i, 5)) Or itemArr(i, 5) = "" Then
            pRate = 0.05
        ElseIf IsNumeric(itemArr(i, 5)) Then
            pRate = CDbl(itemArr(i, 5))
            If pRate > 1 Then pRate = pRate / 100
        Else
            pRate = 0.05
        End If

        ' --- 買い側集計 ---
        Dim buyTotal As Double
        buyTotal = 0
        For j = 2 To UBound(dataArr, 1)
            If CStr(dataArr(j, 4)) = pName Then
                buyTotal = buyTotal + CDbl(dataArr(j, 3))
            End If
        Next j

        ' --- 売り側集計 ---
        Dim sellTotal As Double
        sellTotal = 0
        For j = 2 To UBound(dataArr, 1)
            If CStr(dataArr(j, 6)) = pShort Then
                sellTotal = sellTotal + CDbl(dataArr(j, 3))
            End If
        Next j

        ' --- 取引なしはスキップ（999番は除外） ---
        If buyTotal = 0 And sellTotal = 0 And pNum <> "999" Then GoTo NextParticipant

        ' --- 出力 ---
        With wsResult
            .Cells(outRow, 1).Value = pName
            .Cells(outRow, 2).Value = buyTotal * (1 + pRate)
            .Cells(outRow, 3).Value = buyTotal
            .Cells(outRow, 4).Value = buyTotal * pRate
            .Cells(outRow, 5).Value = pShort
            .Cells(outRow, 6).Value = sellTotal * (1 - pRate)
            .Cells(outRow, 7).Value = sellTotal
            .Cells(outRow, 8).Value = sellTotal * pRate

            ' --- 書式・枠 ---
            Dim col As Long
            For col = 1 To 8
                With .Cells(outRow, col)
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                    If col <> 1 And col <> 5 Then
                        .NumberFormat = "#,##0"
                    End If
                    If col = 2 Or col = 6 Then
                        .Font.Bold = True
                    End If
                End With
            Next col
        End With

        outRow = outRow + 1

NextParticipant:
    Next i
    
    ' --- オートフィルター設定 ---
    With wsResult
        .Range("A9:H9").AutoFilter
    End With

    ' --- サマリー出力 ---
    Dim lastOutRow As Long
    lastOutRow = outRow - 1

    With wsResult
        Dim buySubTotal  As Double
        Dim sellSubTotal As Double
        Dim feeTotal     As Double
        Dim inOutTotal   As Double
        Dim grandTotal   As Double

        buySubTotal = Application.WorksheetFunction.Sum(.Range("B10:B" & lastOutRow))
        sellSubTotal = Application.WorksheetFunction.Sum(.Range("F10:F" & lastOutRow))
        feeTotal = Application.WorksheetFunction.Sum(.Range("D10:D" & lastOutRow)) _
                     + Application.WorksheetFunction.Sum(.Range("H10:H" & lastOutRow))
        inOutTotal = buySubTotal + sellSubTotal
        grandTotal = feeTotal + inOutTotal

        .Cells(2, 2).Value = grandTotal
        .Cells(3, 2).Value = feeTotal
        .Cells(4, 2).Value = inOutTotal
        .Cells(6, 2).Value = buySubTotal
        .Cells(7, 2).Value = sellSubTotal

        ' --- サマリーセルの書式・枠 ---
        Dim summaryRows As Variant
        Dim r As Variant
        summaryRows = Array(2, 3, 4, 6, 7)
        For Each r In summaryRows
            With .Cells(r, 2)
                .NumberFormat = "#,##0"
                .Font.Size = 14
                .Font.Bold = True
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlMedium
            End With
        Next r
        ' B7のみ下辺も太枠
        With .Cells(7, 2)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
        
        ' --- 背景色 ---
        .Range("B2:B4").Interior.Color = RGB(255, 242, 204)
        .Range("B6:B7").Interior.Color = RGB(255, 242, 204)
    End With

    MsgBox "集計が完了しました。", vbInformation

End Sub
