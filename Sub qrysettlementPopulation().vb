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
    Dim temp_str_arr() As String
    index = 0
    For i = 1 To num_cases
        If distribution_tbl.DataBodyRange(i, 1) = qrySettlement_tbl.DataBodyRange(j, 1) Then
            
        End If
        'j=i for the nested for loop because the output from the Needles database is always in alphabetical order and the distribution spreadsheet should also always be in alphabetical order
        For j = i To qry_cases
            If distribution_tbl.DataBodyRange(i, 1) = qrySettlement_tbl.DataBodyRange(j, 1) Then
                'Class
                distribution_tbl.DataBodyRange(i, 2) = qrySettlement_tbl.DataBodyRange(j, 7)
                'Prefix
                distribution_tbl.DataBodyRange(i, 3) = qrySettlement_tbl.DataBodyRange(j, 8)
                'Last Name
                distribution_tbl.DataBodyRange(i, 4) = qrySettlement_tbl.DataBodyRange(j, 9)
                'First Name
                distribution_tbl.DataBodyRange(i, 5) = qrySettlement_tbl.DataBodyRange(j, 10)
                'Role
                distribution_tbl.DataBodyRange(i, 7) = qrySettlement_tbl.DataBodyRange(j, 17)
                'company
                distribution_tbl.DataBodyRange(i, 12) = qrySettlement_tbl.DataBodyRange(j, 18)
                'address
                distribution_tbl.DataBodyRange(i, 13) = qrySettlement_tbl.DataBodyRange(j, 19)
                'address_2
                distribution_tbl.DataBodyRange(i, 14) = qrySettlement_tbl.DataBodyRange(j, 20)
                'city
                distribution_tbl.DataBodyRange(i, 15) = qrySettlement_tbl.DataBodyRange(j, 21)
                'state
                distribution_tbl.DataBodyRange(i, 16) = qrySettlement_tbl.DataBodyRange(j, 22)
                'zipcode
                distribution_tbl.DataBodyRange(i, 17) = qrySettlement_tbl.DataBodyRange(j, 23)
                'Fee A%
                distribution_tbl.DataBodyRange(i, 25) = qrySettlement_tbl.DataBodyRange(j, 30)
                'Fee B%
                distribution_tbl.DataBodyRange(i, 26) = qrySettlement_tbl.DataBodyRange(j, 34)
                'Local Counsel
                distribution_tbl.DataBodyRange(i, 33) = qrySettlement_tbl.DataBodyRange(j, 60)
                'LOCAL CNSL Percent
                distribution_tbl.DataBodyRange(i, 34) = qrySettlement_tbl.DataBodyRange(j, 59)
                'REF FEE 1 Atty
                distribution_tbl.DataBodyRange(i, 38) = qrySettlement_tbl.DataBodyRange(j, 39)
                'REF_FEE_1%
                distribution_tbl.DataBodyRange(i, 39) = qrySettlement_tbl.DataBodyRange(j, 38)
                'REF FEE 2 Atty
                distribution_tbl.DataBodyRange(i, 41) = qrySettlement_tbl.DataBodyRange(j, 44)
                'REF_FEE_2%
                distribution_tbl.DataBodyRange(i, 42) = qrySettlement_tbl.DataBodyRange(j, 43)
                'REF FEE 3 Atty
                distribution_tbl.DataBodyRange(i, 44) = qrySettlement_tbl.DataBodyRange(j, 48)
                'REF_FEE_3%
                distribution_tbl.DataBodyRange(i, 45) = qrySettlement_tbl.DataBodyRange(j, 47)
                'REF FEE 4 Atty
                distribution_tbl.DataBodyRange(i, 47) = qrySettlement_tbl.DataBodyRange(j, 52)
                'REF_FEE_4%
                distribution_tbl.DataBodyRange(i, 48) = qrySettlement_tbl.DataBodyRange(j, 51)
                'DUAL REP CON FEE Atty
                distribution_tbl.DataBodyRange(i, 55) = qrySettlement_tbl.DataBodyRange(j, 127)
                'DUAL REP CON_FEE_Percent
                distribution_tbl.DataBodyRange(i, 56) = qrySettlement_tbl.DataBodyRange(j, 126)
                'LOCAL ATTY w/ EXP
                distribution_tbl.DataBodyRange(i, 64) = qrySettlement_tbl.DataBodyRange(j, 60)
                'LOCAL EXP
                If qrySettlement_tbl.DataBodyRange(j, 149) <> "" Then
                    distribution_tbl.DataBodyRange(i, 65) = qrySettlement_tbl.DataBodyRange(j, 149)
                Else
                    distribution_tbl.DataBodyRange(i, 65) = 0
                End If
                'REF EXP Atty
                temp_str_arr = Split(qrySettlement_tbl.DataBodyRange(j, 71), " - ")
                distribution_tbl.DataBodyRange(i, 66) = temp_str_arr
                'REF EXP
                If qrySettlement_tbl.DataBodyRange(j, 69) <> "" Then
                    distribution_tbl.DataBodyRange(i, 67) = qrySettlement_tbl.DataBodyRange(j, 69)
                Else
                    distribution_tbl.DataBodyRange(i, 67) = 0
                End If
                'CON EXP Atty
                distribution_tbl.DataBodyRange(i, 68) = qrySettlement_tbl.DataBodyRange(j, 127)
                'DUAL REP CON EXP
                If qrySettlement_tbl.DataBodyRange(j, 128) <> "" Then
                    distribution_tbl.DataBodyRange(i, 69) = qrySettlement_tbl.DataBodyRange(j, 128)
                Else
                    distribution_tbl.DataBodyRange(i, 69) = 0
                End If
                'FG EXP
                If qrySettlement_tbl.DataBodyRange(j, 72) <> "" Then
                    distribution_tbl.DataBodyRange(i, 70) = qrySettlement_tbl.DataBodyRange(j, 72)
                Else
                    distribution_tbl.DataBodyRange(i, 70) = 0
                End If
                'MA Case Exp
                If qrySettlement_tbl.DataBodyRange(j, 64) <> "" Then
                    distribution_tbl.DataBodyRange(i, 71) = qrySettlement_tbl.DataBodyRange(j, 64)
                Else
                    distribution_tbl.DataBodyRange(i, 71) = 0
                End If
                
                'Third Party Lien List
                temp_str_arr = Split(qrySettlement_tbl.DataBodyRange(j, 86), " - ")
                distribution_tbl.DataBodyRange(i, 121) = temp_str_arr
                'LOCAL ATTY w/ EXP
                If qrySettlement_tbl.DataBodyRange(j, 84) <> "" Then
                    distribution_tbl.DataBodyRange(i, 122) = qrySettlement_tbl.DataBodyRange(j, 84)
                Else
                    distribution_tbl.DataBodyRange(i, 122) = 0
                End If
                'MA Client Advances/Prior Payments
                If qrySettlement_tbl.DataBodyRange(j, 90) <> "" Then
                    distribution_tbl.DataBodyRange(i, 123) = qrySettlement_tbl.DataBodyRange(j, 90)
                Else
                    distribution_tbl.DataBodyRange(i, 123) = 0
                End If
                'BANKRUPTCY
                If qrySettlement_tbl.DataBodyRange(j, 79) <> "" Then
                    distribution_tbl.DataBodyRange(i, 124) = qrySettlement_tbl.DataBodyRange(j, 79)
                Else
                    distribution_tbl.DataBodyRange(i, 124) = 0
                End If
                'Bankruptcy History
                If qrySettlement_tbl.DataBodyRange(j, 80) > 0 Then
                    distribution_tbl.DataBodyRange(i, 125) = "Y"
                Else
                    distribution_tbl.DataBodyRange(i, 125) = "N"
                End If
                'Total Partial Payments to Client
                If qrySettlement_tbl.DataBodyRange(j, 132) <> "" Then
                    distribution_tbl.DataBodyRange(i, 130) = qrySettlement_tbl.DataBodyRange(j, 132)
                Else
                    distribution_tbl.DataBodyRange(i, 130) = 0
                End If
                'Total Misc. Payments to Client
                If qrySettlement_tbl.DataBodyRange(j, 142) <> "" Then
                    distribution_tbl.DataBodyRange(i, 131) = qrySettlement_tbl.DataBodyRange(j, 142)
                Else
                    distribution_tbl.DataBodyRange(i, 131) = 0
                End If
                'Total Final Payments to Client
                If qrySettlement_tbl.DataBodyRange(j, 137) <> "" Then
                    distribution_tbl.DataBodyRange(i, 132) = qrySettlement_tbl.DataBodyRange(j, 137)
                Else
                    distribution_tbl.DataBodyRange(i, 132) = 0
                End If
                'Total Lien Refund Payments to Client
                If qrySettlement_tbl.DataBodyRange(j, 146) <> "" Then
                    distribution_tbl.DataBodyRange(i, 133) = qrySettlement_tbl.DataBodyRange(j, 146)
                Else
                    distribution_tbl.DataBodyRange(i, 133) = 0
                End If
      
                match = True
                Exit For
            End If
            
     
        Next j
        
        If match = False Then
            distribution_tbl.DataBodyRange(i, 2) = CVErr(xlErrNA)
        End If
        
        
    Next i
End Sub
