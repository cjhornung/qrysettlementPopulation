Sub qrysettlementPopulation()
    Dim distribution_tbl As ListObject
    Set distribution_tbl = Sheets("Distribution").ListObjects("distribution_tbl")
    Dim qrySettlement_tbl As ListObject
    Set qrySettlement_tbl = Sheets("qrySettlement").ListObjects("qrySettlement_tbl")
    Dim num_cases As Integer
    num_cases = distribution_tbl.DataBodyRange.Rows.Count + distribution_tbl.HeaderRowRange.Rows.Count
    qry_cases = qrySettlement_tbl.DataBodyRange.Rows.Count + qrySettlement_tbl.HeaderRowRange.Rows.Count
    Dim match As Boolean
    match = False
    Dim index As Integer
    index = 0
    For i = 1 To num_cases
        If distribution_tbl.DataBodyRange(i, 1) = qrySettlement_tbl.DataBodyRange(j, 1) Then
            
        End If
        For j = 1 To qry_cases
            If distribution_tbl.DataBodyRange(i, 1) = qrySettlement_tbl.DataBodyRange(j, 1) Then
                distribution_tbl.DataBodyRange(i, 2) = qrySettlement_tbl.DataBodyRange(j, 7)
                distribution_tbl.DataBodyRange(i, 3) = qrySettlement_tbl.DataBodyRange(j, 8)
                distribution_tbl.DataBodyRange(i, 4) = qrySettlement_tbl.DataBodyRange(j, 9)
                distribution_tbl.DataBodyRange(i, 5) = qrySettlement_tbl.DataBodyRange(j, 10)
                distribution_tbl.DataBodyRange(i, 7) = qrySettlement_tbl.DataBodyRange(j, 17)
                distribution_tbl.DataBodyRange(i, 12) = qrySettlement_tbl.DataBodyRange(j, 18)
                distribution_tbl.DataBodyRange(i, 13) = qrySettlement_tbl.DataBodyRange(j, 19)
                distribution_tbl.DataBodyRange(i, 14) = qrySettlement_tbl.DataBodyRange(j, 20)
                distribution_tbl.DataBodyRange(i, 15) = qrySettlement_tbl.DataBodyRange(j, 21)
                distribution_tbl.DataBodyRange(i, 16) = qrySettlement_tbl.DataBodyRange(j, 22)
                distribution_tbl.DataBodyRange(i, 17) = qrySettlement_tbl.DataBodyRange(j, 23)
                match = True
            End If
            
     
        Next j
    Next i
End Sub
