'//////////////////////////////////////////////////////////////////////////////////////////////////////
'///////// 데이터 입력
'//////////////////////////////////////////////////////////////////////////////////////////////////////


Module inputOil
    Public inData As CarTotal '임시 옮기는 변수

    Public Sub moveData() '입력한 데이터를 옮기기  Main

        With inData
            .year = Mid(Main.gbt_Date.Text, 1, 4)
            .month = Main.gbtCb_Month.Text
            .ndate = Main.gbt_Date.Text
            .parts = Main.gbtCb_Parts.Text
            .kum = rtnNotComma(Main.gbtTx_Kum.Text)
            .danga = rtnNotComma(Main.gbtTx_Danga.Text)
            .km = rtnNotComma(Main.gbtTx_Km.Text)
            .oil = rtnNotComma(Main.gbtTx_Oil.Text)
            .rkm = rtnNotComma(Main.gbtTx_TKm.Text)
            .km1L = rtnNotComma(Main.gbtTx_KmL.Text)
            .store = Main.gbtTx_Store.Text
            .etc = Main.gbtTx_Etc.Text
        End With
    End Sub

    Public Sub moveTotal() '입력한 데이터를 총괄 결과 데이터로 옮기기  Main
        Dim i As Integer = -1
        Dim mveStr1 As String
        Dim mveStr2 As String

        mveStr1 = rtnNotComma(Main.gbtTx_MTotal.Text)
        mveStr2 = rtnNotComma(Main.gbtTx_TLiter.Text)

        i = Main.gbtCb_Parts.SelectedIndex

        Main.gbtTx_MTotal.Text = CDbl(mveStr1) + CDbl(inData.kum)
        If i = 0 Then
            Main.gbtTx_TLiter.Text = CDbl(mveStr2) + CDbl(inData.oil)

        End If
        Main.gbtTx_TTLKm.Text = Main.gbtTx_Km.Text

        Total1.Kum = Main.gbtTx_MTotal.Text
        Total1.ttlKm = Main.gbtTx_TTLKm.Text
        Total1.avgKm = Main.gbtTx_aveOil.Text
        Total1.wonofLiter = Main.gbtTx_won1L.Text
        Total1.ttlLiter = Main.gbtTx_TLiter.Text

        Main.gbtTx_MTotal.Text = rtnComma(Main.gbtTx_MTotal.Text)
        Main.gbtTx_TTLKm.Text = rtnComma(Main.gbtTx_TTLKm.Text)
        Main.gbtTx_aveOil.Text = rtnComma(Main.gbtTx_aveOil.Text)
        Main.gbtTx_won1L.Text = rtnComma(Main.gbtTx_won1L.Text)
        Main.gbtTx_TLiter.Text = rtnComma(Main.gbtTx_TLiter.Text)
    End Sub

    Public Sub moveOil() '주유 리스트 옮기기  Main
        Dim i As Integer = -1

        iCounter += 1
        ReDim Preserve cartotal1(iCounter)

        With cartotal1(iCounter - 1)
            .year = inData.year
            .month = inData.month
            .ndate = inData.ndate
            .parts = inData.parts
            .kum = inData.kum
            .danga = inData.danga
            .km = inData.km
            .oil = inData.oil
            .rkm = inData.rkm
            .km1L = inData.km1L
            .store = inData.store
            .etc = inData.etc
        End With

        Main.dtGrip_List.RowCount = iCounter + 1
        For i = 0 To iCounter - 1
            With Main.dtGrip_List.Rows(i)
                .Cells(0).Value = cartotal1(i).month
                .Cells(1).Value = cartotal1(i).ndate
                .Cells(2).Value = cartotal1(i).parts
                .Cells(3).Value = rtnComma(cartotal1(i).kum)
                .Cells(4).Value = rtnComma(cartotal1(i).danga)
                .Cells(5).Value = rtnComma(cartotal1(i).km)
                .Cells(6).Value = rtnComma(cartotal1(i).oil)
                .Cells(7).Value = rtnComma(cartotal1(i).rkm)
                .Cells(8).Value = rtnComma(cartotal1(i).km1L)
                .Cells(9).Value = cartotal1(i).store
                .Cells(10).Value = cartotal1(i).etc

                Main.dtGrip_List.CurrentCell = Main.dtGrip_List.Rows(i).Cells(0)
            End With

        Next

    End Sub
    Public Function rtnPartsCal() As String '부품의 실 거리 계산 기능
        Dim rtn As String = "0"
        Dim i As Integer = 0
        Dim gbtNo As Integer = 1

        gbtNo = Main.gbtCb_Parts.SelectedIndex
        Select Case gbtNo
            Case 1 '엔진오일
                rtn = rtnNotComma(Main.gbtTx_Km.Text) - rtnNotComma(Main.gbSub_dtGrip.Rows(0).Cells(9).Value)
            Case 3 '브레이크 패드
                rtn = rtnNotComma(Main.gbtTx_Km.Text) - rtnNotComma(Main.gbSub_dtGrip.Rows(1).Cells(5).Value)
            Case 4 '부동액
                rtn = rtnNotComma(Main.gbtTx_Km.Text) - rtnNotComma(Main.gbSub_dtGrip.Rows(2).Cells(1).Value)
            Case 5 '타이밍 밸트
                rtn = rtnNotComma(Main.gbtTx_Km.Text) - rtnNotComma(Main.gbSub_dtGrip.Rows(0).Cells(1).Value)
            Case 6 '에어콘 가스
                rtn = rtnNotComma(Main.gbtTx_Km.Text) - rtnNotComma(Main.gbSub_dtGrip.Rows(2).Cells(5).Value)
            Case 7 '외주 밸트
                rtn = rtnNotComma(Main.gbtTx_Km.Text) - rtnNotComma(Main.gbSub_dtGrip.Rows(2).Cells(9).Value)
            Case 8 '에어콘필터
                rtn = rtnNotComma(Main.gbtTx_Km.Text) - rtnNotComma(Main.gbSub_dtGrip.Rows(0).Cells(5).Value)
            Case 9 '밧데리
                rtn = rtnNotComma(Main.gbtTx_Km.Text) - rtnNotComma(Main.gbSub_dtGrip.Rows(1).Cells(1).Value)
            Case 10 '타이어
                rtn = rtnNotComma(Main.gbtTx_Km.Text) - rtnNotComma(Main.gbSub_dtGrip.Rows(1).Cells(9).Value)
            Case Else

        End Select

        Main.gbtTx_TKm.Text = rtn

        Return rtn
    End Function


    Public Sub movePart() '부품 교환 리스트 옮기기  Main
        Dim i As Integer = -1

        iParts += 1
        ReDim Preserve Parts1(iParts)

        With Parts1(iParts)
            .year = Mid(Main.gbt_Date.Text, 1, 4)
            .month = Main.gbtCb_Month.Text
            .pdate = Main.gbt_Date.Text
            .part = Main.gbtCb_Parts.Text
            .kum = rtnNotComma(Main.gbtTx_Kum.Text)
            .km = rtnNotComma(Main.gbtTx_Km.Text)
            '.rkm = rtnPartsCal() '
            .rkm = rtnNotComma(Main.gbtTx_TKm.Text)
            .store = Main.gbtTx_Store.Text
            .etc = Main.gbtTx_Etc.Text
        End With

        Main.gbGridStore.RowCount = iParts + 1
        For i = 0 To iParts

            With Main.gbGridStore
                .Rows(i).Cells(0).Value = Parts1(i).month
                .Rows(i).Cells(1).Value = Parts1(i).pdate
                .Rows(i).Cells(2).Value = Parts1(i).part
                .Rows(i).Cells(3).Value = rtnComma(Parts1(i).kum)

                .Rows(i).Cells(4).Value = rtnComma(Parts1(i).km)
                .Rows(i).Cells(5).Value = rtnComma(Parts1(i).rkm)
                .Rows(i).Cells(6).Value = Parts1(i).store
                .Rows(i).Cells(7).Value = Parts1(i).etc

            End With

            Main.gbGridStore.CurrentCell = Main.gbGridStore.Rows(i).Cells(0)
        Next
        Main.gbtTx_MTotal.Text = CDbl(rtnNotComma(Main.gbtTx_MTotal.Text)) + CDbl(rtnNotComma(Main.gbtTx_Kum.Text))
    End Sub
    Public Sub movePartsLen()
        '부품 교환후 이동 거리
        Dim mveStr1 As String
        Dim mveStr2 As String

        mveStr1 = rtnNotComma(Main.gbtTx_Km.Text)

       
        mveStr2 = rtnNotComma(Main.gbSub_dtGrip.Rows(0).Cells(1).Value)
        Main.gbSub_dtGrip.Rows(0).Cells(2).Value = rtnComma(mveStr1 - mveStr2)
        'Main.gbSub_dtGrip.Rows(0).Cells(2).Value = rtnComma(Main.gbtTx_TKm.Text)

        mveStr2 = rtnNotComma(Main.gbSub_dtGrip.Rows(0).Cells(5).Value)
        Main.gbSub_dtGrip.Rows(0).Cells(6).Value = rtnComma(mveStr1 - mveStr2)
        'Main.gbSub_dtGrip.Rows(0).Cells(6).Value = rtnComma(Main.gbtTx_TKm.Text)

        mveStr2 = rtnNotComma(Main.gbSub_dtGrip.Rows(0).Cells(9).Value)
        Main.gbSub_dtGrip.Rows(0).Cells(10).Value = rtnComma(mveStr1 - mveStr2)
        'Main.gbSub_dtGrip.Rows(0).Cells(10).Value = rtnComma(Main.gbtTx_TKm.Text)


        mveStr2 = rtnNotComma(Main.gbSub_dtGrip.Rows(1).Cells(1).Value)
        Main.gbSub_dtGrip.Rows(1).Cells(2).Value = rtnComma(mveStr1 - mveStr2)
        'Main.gbSub_dtGrip.Rows(1).Cells(2).Value = rtnComma(Main.gbSub_dtGrip.Rows(1).Cells(2).Value)

        'mveStr1 = rtnNotComma(Main.gbtTx_Km.Text)
        mveStr2 = rtnNotComma(Main.gbSub_dtGrip.Rows(1).Cells(5).Value)
        Main.gbSub_dtGrip.Rows(1).Cells(6).Value = rtnComma(mveStr1 - mveStr2)
        'Main.gbSub_dtGrip.Rows(1).Cells(6).Value = rtnComma(Main.gbtTx_TKm.Text)

        mveStr2 = rtnNotComma(Main.gbSub_dtGrip.Rows(1).Cells(9).Value)
        Main.gbSub_dtGrip.Rows(1).Cells(10).Value = rtnComma(mveStr1 - mveStr2)
        'Main.gbSub_dtGrip.Rows(1).Cells(10).Value = rtnComma(Main.gbSub_dtGrip.Rows(1).Cells(10).Value)


        mveStr2 = rtnNotComma(Main.gbSub_dtGrip.Rows(2).Cells(1).Value)
        Main.gbSub_dtGrip.Rows(2).Cells(2).Value = rtnComma(mveStr1 - mveStr2)
        'Main.gbSub_dtGrip.Rows(2).Cells(2).Value = rtnComma(Main.gbtTx_TKm.Text)

        mveStr2 = rtnNotComma(Main.gbSub_dtGrip.Rows(2).Cells(5).Value)
        Main.gbSub_dtGrip.Rows(2).Cells(6).Value = rtnComma(mveStr1 - mveStr2)
        'Main.gbSub_dtGrip.Rows(2).Cells(6).Value = rtnComma(Main.gbtTx_TKm.Text)

        mveStr2 = rtnNotComma(Main.gbSub_dtGrip.Rows(2).Cells(9).Value)
        Main.gbSub_dtGrip.Rows(2).Cells(10).Value = rtnComma(mveStr1 - mveStr2)
        'Main.gbSub_dtGrip.Rows(2).Cells(10).Value = rtnComma(Main.gbtTx_TKm.Text)

      

   

       
    End Sub
    Public Sub moveParts() '입력한 데이터를 교환 부품 리스트 옮기기  Main
        Dim i As Integer = -1
        Dim j As Integer = -1
        Dim str As String = Main.gbtTx_Km.Text
        Dim strDate As String = Main.gbt_Date.Text
        i = Main.gbtCb_Parts.SelectedIndex

        Select Case i
            Case 1 '엔진오일
                'Main.gbSub_dtGrip.Rows(0).Cells(8).Value = 1 '엔진 오일
                Main.gbSub_dtGrip.Rows(0).Cells(9).Value = str
                'Main.gbSub_dtGrip.Rows(0).Cells(10).Value = Main.gbtTx_Km.Text - Main.gbSub_dtGrip.Rows(0).Cells(9).Value
                Main.gbSub_dtGrip.Rows(0).Cells(11).Value = strDate
            Case 2 '브레이크패드
                'Main.gbSub_dtGrip.Rows(1).Cells(4).Value = 1 '브레이크 패드
                Main.gbSub_dtGrip.Rows(1).Cells(5).Value = str
                'Main.gbSub_dtGrip.Rows(1).Cells(6).Value = Main.gbtTx_Km.Text - Main.gbSub_dtGrip.Rows(1).Cells(5).Value
                Main.gbSub_dtGrip.Rows(1).Cells(7).Value = strDate
            Case 3 '부동액
                'Main.gbSub_dtGrip.Rows(2).Cells(0).Value = 1 '부동액
                Main.gbSub_dtGrip.Rows(2).Cells(1).Value = str
                'Main.gbSub_dtGrip.Rows(2).Cells(2).Value = Main.gbtTx_Km.Text - Main.gbSub_dtGrip.Rows(2).Cells(1).Value
                Main.gbSub_dtGrip.Rows(2).Cells(3).Value = strDate
            Case 4 '타임 밸트
                'Main.gbSub_dtGrip.Rows(0).Cells(0).Value = 1 '타임 밸트
                Main.gbSub_dtGrip.Rows(0).Cells(1).Value = str
                'Main.gbSub_dtGrip.Rows(0).Cells(2).Value = Main.gbtTx_Km.Text - Main.gbSub_dtGrip.Rows(0).Cells(1).Value
                Main.gbSub_dtGrip.Rows(0).Cells(3).Value = strDate
            Case 5 '에어컨가스
                'Main.gbSub_dtGrip.Rows(2).Cells(4).Value = 1 '에어컨 가스
                Main.gbSub_dtGrip.Rows(2).Cells(5).Value = str
                'Main.gbSub_dtGrip.Rows(2).Cells(6).Value = Main.gbtTx_Km.Text - Main.gbSub_dtGrip.Rows(2).Cells(5).Value
                Main.gbSub_dtGrip.Rows(2).Cells(7).Value = strDate
            Case 6 '외부 밸트
                'Main.gbSub_dtGrip.Rows(2).Cells(8).Value = 1 '외부 밸트
                Main.gbSub_dtGrip.Rows(2).Cells(9).Value = str
                'Main.gbSub_dtGrip.Rows(2).Cells(10).Value = Main.gbtTx_Km.Text - Main.gbSub_dtGrip.Rows(2).Cells(9).Value
                Main.gbSub_dtGrip.Rows(2).Cells(11).Value = strDate
            Case 7 '에어컨 필터
                'Main.gbSub_dtGrip.Rows(0).Cells(4).Value = 1 '에어컨 필터
                Main.gbSub_dtGrip.Rows(0).Cells(5).Value = str
                'Main.gbSub_dtGrip.Rows(0).Cells(6).Value = Main.gbtTx_Km.Text - Main.gbSub_dtGrip.Rows(0).Cells(5).Value
                Main.gbSub_dtGrip.Rows(0).Cells(7).Value = strDate
            Case 8 '밧데리
                'Main.gbSub_dtGrip.Rows(1).Cells(0).Value = 1 '밧데리
                Main.gbSub_dtGrip.Rows(1).Cells(1).Value = str
                'Main.gbSub_dtGrip.Rows(1).Cells(2).Value = Main.gbtTx_Km.Text - Main.gbSub_dtGrip.Rows(1).Cells(1).Value
                Main.gbSub_dtGrip.Rows(1).Cells(3).Value = strDate
            Case 9 '타이어
                'Main.gbSub_dtGrip.Rows(1).Cells(8).Value = 1 '타이어
                Main.gbSub_dtGrip.Rows(1).Cells(9).Value = str
                'Main.gbSub_dtGrip.Rows(1).Cells(10).Value = Main.gbtTx_Km.Text - Main.gbSub_dtGrip.Rows(1).Cells(9).Value
                Main.gbSub_dtGrip.Rows(1).Cells(11).Value = strDate
            Case 10 '냉각수밸브

            Case 11 '쇼바

            Case 12 '가스켓

            Case 13 '등속조인트


            Case Else


        End Select

        For j = 0 To 2

            With repairParts1(j)
                .name = Main.gbSub_dtGrip.Rows(j).Cells(0).Value
                .chgKm = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(1).Value)
                .move = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(2).Value)
                .pdate = Main.gbSub_dtGrip.Rows(j).Cells(3).Value

                .name1 = Main.gbSub_dtGrip.Rows(j).Cells(4).Value
                .chgKm1 = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(5).Value)
                .move1 = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(6).Value)
                .pdate1 = Main.gbSub_dtGrip.Rows(j).Cells(7).Value

                .name2 = Main.gbSub_dtGrip.Rows(j).Cells(8).Value
                .chgKm2 = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(9).Value)
                .move2 = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(10).Value)
                .pdate2 = Main.gbSub_dtGrip.Rows(j).Cells(11).Value

                Main.gbSub_dtGrip.Rows(j).Cells(1).Value = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(1).Value)
                Main.gbSub_dtGrip.Rows(j).Cells(2).Value = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(2).Value)

                Main.gbSub_dtGrip.Rows(j).Cells(5).Value = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(5).Value)
                Main.gbSub_dtGrip.Rows(j).Cells(6).Value = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(6).Value)

                Main.gbSub_dtGrip.Rows(j).Cells(9).Value = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(9).Value)
                Main.gbSub_dtGrip.Rows(j).Cells(10).Value = rtnNotComma(Main.gbSub_dtGrip.Rows(j).Cells(10).Value)
            End With
        Next


    End Sub

    Public Sub moveStore()
        '주유소 주소 리스트 입력  (form : subStoreList)
        Dim i As Integer = -1

        iStoreCounter += 1

        ReDim Preserve subStore1(iStoreCounter)

        With subStore1(iStoreCounter)
            .store = subStoreList.txStore.Text
            .etc = subStoreList.txEtc.Text
        End With

        subStoreList.subGrid.RowCount = iStoreCounter + 1
        For i = 0 To iStoreCounter
            With subStoreList.subGrid
                .Rows(i).Cells(0).Value = subStore1(i).store
                .Rows(i).Cells(1).Value = subStore1(i).etc

            End With
        Next
    End Sub

    Public Sub movePartsAddress()
        '수리점 과 주소를 옮기기 (form : subPartsList)
        Dim i As Integer = -1

        iPartsStore += 1

        ReDim Preserve PartAddressList(iPartsStore)

        With PartAddressList(iPartsStore)
            .store = subPartsList.txStore.Text
            .etc = subPartsList.txEtc.Text

        End With

        subPartsList.subGrid.RowCount = iPartsStore + 1
        For i = 0 To iPartsStore
            With subPartsList.subGrid.Rows(i)
                .Cells(0).Value = PartAddressList(i).store
                .Cells(1).Value = PartAddressList(i).etc
            End With
        Next
    End Sub

End Module
