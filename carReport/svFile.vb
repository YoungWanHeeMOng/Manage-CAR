'//////////////////////////////////////////////////////////////////////////////////////////////////////
'///////// 파일  저장
'//////////////////////////////////////////////////////////////////////////////////////////////////////



Module svFile
    Public svspc As String = "♣"  '구분 기호
    'Public str As String = vbNullString '읽은 소스

    Public svFolder As String = dataFolder   '데이터 폴더

    'Public cartotal1() As CarTotal '데이터를 임시 저장하는 곳 , 원본 소스 데이터
    'Public subStore1() As subStore '주유소와 주소를 임시 저장하는 곳



    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 주유 내역 리스트 mainsource data  Reading
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub svLstFile()
        '메인 데이터 소스 읽어 들이기
        Dim iSave As Integer = -1

        str = ""

        Using sv As StreamWriter = New StreamWriter(svFolder & "TotalList.txt", False, System.Text.Encoding.UTF8)
            str = "년	월	날짜	부품내용	 금액 	 단가 	 주행거리 	경유량( liter)	 실거리Km 	 ?Km/1L 	주유소	비고"
            sv.WriteLine(str)
            For iSave = 0 To iCounter - 1
                'With cartotal1(iSave)
                '    str = .year
                '    str = str & svspc & .month
                '    str = str & svspc & .ndate
                '    str = str & svspc & .parts
                '    str = str & svspc & .kum
                '    str = str & svspc & .danga
                '    str = str & svspc & .km
                '    str = str & svspc & .oil
                '    str = str & svspc & .rkm
                '    str = str & svspc & .km1L
                '    str = str & svspc & .store
                '    str = str & svspc & .etc
                'End With
                With Main.dtGrip_List.Rows(iSave)
                    str = cartotal1(iSave).year
                    str = str & svspc & .Cells(0).Value
                    str = str & svspc & .Cells(1).Value
                    str = str & svspc & .Cells(2).Value
                    str = str & svspc & rtnNotComma(.Cells(3).Value)
                    str = str & svspc & rtnNotComma(.Cells(4).Value)
                    str = str & svspc & rtnNotComma(.Cells(5).Value)
                    str = str & svspc & rtnNotComma(.Cells(6).Value)
                    str = str & svspc & rtnNotComma(.Cells(7).Value)
                    str = str & svspc & rtnNotComma(.Cells(8).Value)
                    str = str & svspc & .Cells(9).Value
                    str = str & svspc & .Cells(10).Value
                End With

                If str.Length > 30 Then sv.WriteLine(str)
                str = ""
            Next

        End Using


        'FileOpen(1, svFolder & "TotalList.txt", OpenMode.Output, OpenAccess.Write)

        'PrintLine(1, "년	월	날짜	부품내용	 금액 	 단가 	 주행거리 	경유량( liter)	 실거리Km 	 ?Km/1L 	주유소	비고")

        'For iSave = 0 To iCounter
        '    'With cartotal1(iSave)
        '    '    str = .year
        '    '    str = str & svspc & .month
        '    '    str = str & svspc & .ndate
        '    '    str = str & svspc & .parts
        '    '    str = str & svspc & .kum
        '    '    str = str & svspc & .danga
        '    '    str = str & svspc & .km
        '    '    str = str & svspc & .oil
        '    '    str = str & svspc & .rkm
        '    '    str = str & svspc & .km1L
        '    '    str = str & svspc & .store
        '    '    str = str & svspc & .etc
        '    'End With
        '    With Main.dtGrip_List.Rows(iSave)
        '        str = cartotal1(iSave).year
        '        str = str & svspc & .Cells(0).Value
        '        str = str & svspc & .Cells(1).Value
        '        str = str & svspc & .Cells(2).Value
        '        str = str & svspc & rtnNotComma(.Cells(3).Value)
        '        str = str & svspc & rtnNotComma(.Cells(4).Value)
        '        str = str & svspc & rtnNotComma(.Cells(5).Value)
        '        str = str & svspc & rtnNotComma(.Cells(6).Value)
        '        str = str & svspc & rtnNotComma(.Cells(7).Value)
        '        str = str & svspc & rtnNotComma(.Cells(8).Value)
        '        str = str & svspc & .Cells(9).Value
        '        str = str & svspc & .Cells(10).Value
        '    End With
        '    PrintLine(1, str)
        '    str = ""
        'Next
        FileClose(1)
    End Sub

   

    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 주유소 리스트  Save Store and Address
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub svStoreFile()
        '메인 데이터 소스 읽어 들이기
        Dim iSave As Integer = -1
        str = ""

        Using sv As StreamWriter = New StreamWriter(svFolder & "StoreList.txt", False, System.Text.Encoding.UTF8)
            str = "주유소	주소"
            sv.WriteLine(str)
            For iSave = 0 To iStoreCounter
                With subStore1(iSave)
                    str = .store
                    str = str & svspc & .etc

                End With
                If str.Length > 10 Then sv.WriteLine(str)
                str = ""
            Next

        End Using

        'FileOpen(1, svFolder & "StoreList.txt", OpenMode.Output, OpenAccess.Write)
        'PrintLine(1, "주유소	주소")

        'For iSave = 0 To iStoreCounter
        '    With subStore1(iSave)
        '        str = .store
        '        str = str & svspc & .etc

        '    End With
        '    PrintLine(1, str)
        '    str = ""
        'Next

        FileClose()
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 정비소 리스트  Save Store and Address
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub svRepairFile()
        '메인 데이터 소스 읽어 들이기
        Dim iSave As Integer = -1
        str = ""

        Using sv As StreamWriter = New StreamWriter(svFolder & "SPartsStoreList.txt", False, System.Text.Encoding.UTF8)
            str = " 정비소	주소"
            sv.WriteLine(str)
            For iSave = 0 To iPartsStore
                With PartAddressList(iSave)
                    str = .store
                    str = str & svspc & .etc

                End With
                If str.Length > 20 Then sv.WriteLine(str)
                str = ""
            Next


        End Using
        'FileOpen(1, svFolder & "SPartsStoreList.txt", OpenMode.Output, OpenAccess.Write)
        'PrintLine(1, " 정비소	주소")

        'For iSave = 0 To iPartsStore
        '    With PartAddressList(iSave)
        '        str = .store
        '        str = str & svspc & .etc

        '    End With
        '    PrintLine(1, str)
        '    str = ""
        'Next

        FileClose()
    End Sub
    
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 부품 수리 내역 리스트
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    'Public Parts1() As Parts

    Public Sub svParts()
        '메인 데이터 소스 읽어 들이기
        Dim iSave As Integer = -1
        str = ""

        Using sv As StreamWriter = New StreamWriter(svFolder & "SPartsStore.txt", False, System.Text.Encoding.UTF8)
            str = "년	월	날짜	부품내용	 금액 	 주행거리 	 실거리Km 	정비소	비고"

            sv.WriteLine(str)
            For iSave = 0 To iParts - 1
                'With Parts1(iSave)
                '    str = .year
                '    str = str & svspc & .month
                '    str = str & svspc & .pdate
                '    str = str & svspc & .part
                '    str = str & svspc & .kum
                '    str = str & svspc & .km
                '    str = str & svspc & .rkm
                '    str = str & svspc & .store
                '    str = str & svspc & .etc
                'End With

                With Main.gbGridStore.Rows(iSave)
                    str = Parts1(iSave).year
                    str = str & svspc & .Cells(0).Value
                    str = str & svspc & .Cells(1).Value
                    str = str & svspc & .Cells(2).Value
                    str = str & svspc & rtnNotComma(.Cells(3).Value)
                    str = str & svspc & rtnNotComma(.Cells(4).Value)
                    str = str & svspc & rtnNotComma(.Cells(5).Value)
                    str = str & svspc & .Cells(6).Value
                    str = str & svspc & .Cells(7).Value
                    'str = str & svspc & .Cells(8).Value
                End With
                If str.Length > 15 Then sv.WriteLine(str)
                str = ""
            Next

        End Using
        'FileOpen(1, svFolder & "SPartsStore.txt", OpenMode.Output, OpenAccess.Write)
        'PrintLine(1, "년	월	날짜	부품내용	 금액 	 주행거리 	 실거리Km 	정비소	비고")

        'For iSave = 0 To iParts
        '    'With Parts1(iSave)
        '    '    str = .year
        '    '    str = str & svspc & .month
        '    '    str = str & svspc & .pdate
        '    '    str = str & svspc & .part
        '    '    str = str & svspc & .kum
        '    '    str = str & svspc & .km
        '    '    str = str & svspc & .rkm
        '    '    str = str & svspc & .store
        '    '    str = str & svspc & .etc
        '    'End With

        '    With Main.gbGridStore.Rows(iSave)
        '        str = Parts1(iSave).year
        '        str = str & svspc & .Cells(0).Value
        '        str = str & svspc & .Cells(1).Value
        '        str = str & svspc & .Cells(2).Value
        '        str = str & svspc & rtnNotComma(.Cells(3).Value)
        '        str = str & svspc & rtnNotComma(.Cells(4).Value)
        '        str = str & svspc & rtnNotComma(.Cells(5).Value)
        '        str = str & svspc & .Cells(6).Value
        '        str = str & svspc & .Cells(7).Value
        '        'str = str & svspc & .Cells(8).Value
        '    End With
        '    PrintLine(1, str)
        '    str = ""
        'Next
        FileClose(1)
    End Sub
   

    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 부품 종류  교환 리스트
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    'Public repairParts1() As repairParts
    Public cntParts As Integer = -1
    Public Sub svRepairPartsList()
        '메인 데이터 소스 읽어 들이기
        Dim iSave As Integer = -1
        str = ""
        Using sv As StreamWriter = New StreamWriter(svFolder & "Spartslist.txt", False, System.Text.Encoding.UTF8)
            str = "부품이름,교환거리,이동거리,교환일자,  부품이름,교환거리,이동거리,교환일자,  부품이름,교환거리,이동거리,교환일자"
            sv.WriteLine(str)
            For iSave = 0 To 2
                'With repairParts1(iSave)
                '    str = .name
                '    str = str & svspc & .chgKm
                '    str = str & svspc & .move
                '    str = str & svspc & .pdate

                '    str = str & svspc & .name1
                '    str = str & svspc & .chgKm1
                '    str = str & svspc & .move1
                '    str = str & svspc & .pdate1

                '    str = str & svspc & .name2
                '    str = str & svspc & .chgKm2
                '    str = str & svspc & .move2
                '    str = str & svspc & .pdate2
                'End With
                With Main.gbSub_dtGrip.Rows(iSave)
                    str = .Cells(0).Value
                    str = str & svspc & rtnNotComma(.Cells(1).Value)
                    str = str & svspc & rtnNotComma(.Cells(2).Value)
                    str = str & svspc & .Cells(3).Value

                    str = str & svspc & .Cells(4).Value
                    str = str & svspc & rtnNotComma(.Cells(5).Value)
                    str = str & svspc & rtnNotComma(.Cells(6).Value)
                    str = str & svspc & .Cells(7).Value

                    str = str & svspc & .Cells(8).Value
                    str = str & svspc & rtnNotComma(.Cells(9).Value)
                    str = str & svspc & rtnNotComma(.Cells(10).Value)
                    str = str & svspc & .Cells(11).Value
                End With
                If str.Length > 15 Then sv.WriteLine(str)
                str = ""
            Next

        End Using
        'FileOpen(1, svFolder & "Spartslist.txt", OpenMode.Output, OpenAccess.Write)
        'PrintLine(1, "부품이름,교환거리,이동거리,교환일자,  부품이름,교환거리,이동거리,교환일자,  부품이름,교환거리,이동거리,교환일자")

        'For iSave = 0 To 2
        '    'With repairParts1(iSave)
        '    '    str = .name
        '    '    str = str & svspc & .chgKm
        '    '    str = str & svspc & .move
        '    '    str = str & svspc & .pdate

        '    '    str = str & svspc & .name1
        '    '    str = str & svspc & .chgKm1
        '    '    str = str & svspc & .move1
        '    '    str = str & svspc & .pdate1

        '    '    str = str & svspc & .name2
        '    '    str = str & svspc & .chgKm2
        '    '    str = str & svspc & .move2
        '    '    str = str & svspc & .pdate2
        '    'End With
        '    With Main.gbSub_dtGrip.Rows(iSave)
        '        str = .Cells(0).Value
        '        str = str & svspc & rtnNotComma(.Cells(1).Value)
        '        str = str & svspc & rtnNotComma(.Cells(2).Value)
        '        str = str & svspc & .Cells(3).Value

        '        str = str & svspc & .Cells(4).Value
        '        str = str & svspc & rtnNotComma(.Cells(5).Value)
        '        str = str & svspc & rtnNotComma(.Cells(6).Value)
        '        str = str & svspc & .Cells(7).Value

        '        str = str & svspc & .Cells(8).Value
        '        str = str & svspc & rtnNotComma(.Cells(9).Value)
        '        str = str & svspc & rtnNotComma(.Cells(10).Value)
        '        str = str & svspc & .Cells(11).Value
        '    End With
        '    PrintLine(1, str)
        '    str = ""
        'Next
        FileClose()
    End Sub
    

    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 총 결과  리스트
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    'Public Total1 As Total

    Public Sub svTotalCumsumer()
        '메인 데이터 소스 읽어 들이기
        str = ""

        Using sv As StreamWriter = New StreamWriter(svFolder & "Stotalcar.txt", False, System.Text.Encoding.UTF8)
            str = "총금액,총주유,총거리,편균연비,리터당금액"
            sv.WriteLine(str)
            With Main
                str = rtnNotComma(.gbtTx_MTotal.Text)
                str = str & svspc & rtnNotComma(.gbtTx_TLiter.Text)
                str = str & svspc & rtnNotComma(.gbtTx_TTLKm.Text)
                str = str & svspc & rtnNotComma(.gbtTx_aveOil.Text)
                str = str & svspc & rtnNotComma(.gbtTx_won1L.Text)

            End With
            sv.WriteLine(str)
            str = ""

        End Using
        'FileOpen(1, svFolder & "Stotalcar.txt", OpenMode.Output, OpenAccess.Write)
        'PrintLine(1, "총금액,총주유,총거리,편균연비,리터당금액")

        ''With Total1
        ''    str = .Kum
        ''    str = str & svspc & .ttlLiter
        ''    str = str & svspc & .ttlKm
        ''    str = str & svspc & .avgKm
        ''    str = str & svspc & .wonofLiter

        ''End With
        'With Main
        '    str = rtnNotComma(.gbtTx_MTotal.Text)
        '    str = str & svspc & rtnNotComma(.gbtTx_TLiter.Text)
        '    str = str & svspc & rtnNotComma(.gbtTx_TTLKm.Text)
        '    str = str & svspc & rtnNotComma(.gbtTx_aveOil.Text)
        '    str = str & svspc & rtnNotComma(.gbtTx_won1L.Text)

        'End With
        'PrintLine(1, str)
        'str = ""
        FileClose()
    End Sub
    

End Module
