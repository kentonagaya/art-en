Option Explicit

' --- 相互売買警告機能 ---
Public Function IsMutualTradeWarning(ByVal buyerNo As String, ByVal sellerNo As String) As Boolean

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim aNo As String, dNo As String

    Set ws = Worksheets("相互売買警告リスト")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow   ' 1行目は見出し想定

        aNo = Trim(ws.Cells(i, "A").Value)
        dNo = Trim(ws.Cells(i, "D").Value)

        '--- どちらか空欄なら無効 ---
        If aNo = "" Or dNo = "" Then
            GoTo ContinueLoop
        End If

        '--- 組み合わせチェック（順不同） ---
        If (buyerNo = aNo And sellerNo = dNo) _
        Or (buyerNo = dNo And sellerNo = aNo) Then
            IsMutualTradeWarning = True
            Exit Function
        End If

ContinueLoop:
    Next i

    IsMutualTradeWarning = False

End Function


