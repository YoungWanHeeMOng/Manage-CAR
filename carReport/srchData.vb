'//////////////////////////////////////////////////////////////////////////////////////////////////////
'///////// 데이터 찾기 
'//////////////////////////////////////////////////////////////////////////////////////////////////////



Module srchData

    Public Sub srchStore() '주유소 찾기
        Dim i As Integer = -1 '임시 카운터
        Dim j As Integer = -1 '찾은 주유소 임시 카운터

        Dim strStore As String '원 주유소 임시 
        Dim strStoreSrch As String '찾은 주유소 임시
        Dim subStoreSrch() As subStore '주유소와 주소를 임시 저장하는 곳

        strStoreSrch = Replace(subStoreList.txStore.Text, " ", "")

        If strStoreSrch = "" Then
            subStoreList.subGrid.Visible = True
            subStoreList.subGrid1.Visible = False
            Exit Sub
        End If

        For i = 0 To iStoreCounter  '찾기
            strStore = subStore1(i).store
            If strStore Like "*" & strStoreSrch & "*" Then
                j += 1

                ReDim Preserve subStoreSrch(j + 1)
                With subStoreSrch(j)
                    .store = subStoreList.subGrid.Rows(i).Cells(0).Value
                    .etc = subStoreList.subGrid.Rows(i).Cells(1).Value
                End With
            End If
        Next

        If j >= 0 Then
            subStoreList.subGrid1.Visible = True
            subStoreList.subGrid.Visible = False

            subStoreList.subGrid1.RowCount = j + 1
            For i = 0 To j '넣기
                With subStoreList.subGrid1
                    .Rows(i).Cells(0).Value = subStoreSrch(i).store
                    .Rows(i).Cells(1).Value = subStoreSrch(i).etc
                End With
            Next
        End If

    End Sub

    Public Sub srchPartsStore() '수리점 찾기
        Dim i As Integer = -1 '임시 카운터
        Dim j As Integer = -1 '찾은 주유소 임시 카운터

        Dim strStore As String '원 주유소 임시 
        Dim strStoreSrch As String = "" '찾은 주유소 임시
        Dim subPartsSrch() As subStore '주유소와 주소를 임시 저장하는 곳

        strStoreSrch = Replace(subPartsList.txStore.Text, " ", "")

        If strStoreSrch = "" Then
            subPartsList.subGrid.Visible = True
            subPartsList.subGrid1.Visible = False
            Exit Sub
        End If

        For i = 0 To iPartsStore  '찾기
            strStore = PartAddressList(i).store
            If strStore Like "*" & strStoreSrch & "*" Then
                j += 1

                ReDim Preserve subPartsSrch(j + 1)
                With subPartsSrch(j)
                    .store = subPartsList.subGrid.Rows(i).Cells(0).Value
                    .etc = subPartsList.subGrid.Rows(i).Cells(1).Value
                End With
            End If
        Next

        If j >= 0 Then
            subPartsList.subGrid1.Visible = True
            subPartsList.subGrid.Visible = False

            subPartsList.subGrid1.RowCount = j + 1
            For i = 0 To j '넣기
                With subPartsList.subGrid1
                    .Rows(i).Cells(0).Value = subPartsSrch(i).store
                    .Rows(i).Cells(1).Value = subPartsSrch(i).etc
                End With
            Next
        End If

    End Sub
End Module
