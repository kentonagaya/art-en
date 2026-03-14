Option Explicit

Sub 集計()

    Sheets("集計").Range("A5:I999").Clear
    
    Dim dicBuyer As Object, dicSeller As Object
    Set dicBuyer = CreateObject("Scripting.Dictionary")
    Set dicSeller = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim data As Variant, buyerKeys As Variant, sellerKeys As Variant
    
    data = Sheets("伝票一覧").Range("A1").CurrentRegion
    
    ' 買主データの集計
    
    For i = 2 To UBound(data)
    
        If Not dicBuyer.Exists(data(i, 4)) Then
            dicBuyer.Add data(i, 4), data(i, 3)
        Else
            dicBuyer(data(i, 4)) = dicBuyer(data(i, 4)) + data(i, 3)
        End If
    
    Next
    
    ' 売主データの集計
    
    For i = 2 To UBound(data)
    
        If Not dicSeller.Exists(data(i, 6)) Then
            dicSeller.Add data(i, 6), data(i, 3)
        Else
            dicSeller(data(i, 6)) = dicSeller(data(i, 6)) + data(i, 3)
        End If
        
    Next
    
    With Sheets("集計")
    
        ' 買主データの出力
        buyerKeys = dicBuyer.keys
        For i = LBound(buyerKeys) To UBound(buyerKeys)
            .Cells(i + 5, 1) = buyerKeys(i) ' 買主
            .Cells(i + 5, 2) = dicBuyer(buyerKeys(i)) * 1.05 ' 合計落札額と手数料5％
            .Cells(i + 5, 3) = dicBuyer(buyerKeys(i)) ' 合計落札額
            .Cells(i + 5, 4) = dicBuyer(buyerKeys(i)) * 0.05 ' 合計手数料
        Next
        
        ' 売主データの出力
        sellerKeys = dicSeller.keys
        For i = LBound(sellerKeys) To UBound(sellerKeys)
            .Cells(i + 5, 5) = sellerKeys(i) ' 売主
            .Cells(i + 5, 7) = dicSeller(sellerKeys(i)) ' 手数料込金額
            .Cells(i + 5, 6) = dicSeller(sellerKeys(i)) * 0.95 ' 手数料抜金額
            .Cells(i + 5, 8) = dicSeller(sellerKeys(i)) * 0.05 ' 手数料のみの金額
        Next
        
    End With

End Sub

