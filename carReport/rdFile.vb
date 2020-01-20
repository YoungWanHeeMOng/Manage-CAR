'//////////////////////////////////////////////////////////////////////////////////////////////////////
'///////// 파일 읽기 
'//////////////////////////////////////////////////////////////////////////////////////////////////////
Module rdFile
    Public spc As String = "♣"  '구분 기호
    Public str As String = vbNullString '읽은 소스
    Public strSplt() As String '소스 구분

    Public dataFolder As String = "D:\Back up Data\MyCar_Data\" '데이터 폴더

    Public cartotal1() As CarTotal '데이터를 임시 저장하는 곳 , 원본 소스 데이터
    Public subStore1() As subStore '주유소와 주소를 임시 저장하는 곳

    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 주유 경유 총 리스트 mainsource data  Reading
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub readFile()
        '메인 데이터 소스 읽어 들이기
        'FileOpen(1, dataFolder & "\data\TotalList.txt", OpenMode.Input, OpenAccess.Read)
        'FileOpen(1, dataFolder & "TotalList.txt", OpenMode.Input, OpenAccess.Read)
        'str = LineInput(1)
        'Do Until EOF(1)

        '    iCounter += 1
        '    str = LineInput(1)

        '    Main.t.Text = str

        '    'Main.t.Font = New System.Drawing.Font("굴림", 20, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        '    str = Main.t.Text

        '    sltData(str)

        'Loop


        Using sr As StreamReader = New StreamReader(dataFolder & "TotalList.txt", System.Text.Encoding.UTF8)
            str = sr.ReadLine

            Do While str <> Nothing
                iCounter += 1
                str = sr.ReadLine

                'Main.t.Text = str

                'Main.t.Font = New System.Drawing.Font("굴림", 20, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
                'str = Main.t.Text

                If str = Nothing Then
                    'MsgBox("없다")
                    iCounter -= 1
                    GoTo RdEnd
                End If

                If Len(str) > 30 Then
                    sltData(str)
                Else
                    iCounter -= 1
                End If

            Loop

RdEnd:
        End Using

        'ReDim Preserve cartotal1(iCounter - 1)   '2018-07-27 새로 삽입

        sltGrid()

        FileClose()
    End Sub

    Public Sub sltData(ByRef strDt As String)
        '파일 데이터 읽어 분리하는 루틴
        ReDim Preserve cartotal1(iCounter)

        strSplt = Split(strDt, spc)

        With cartotal1(iCounter - 1)
            .year = strSplt(0)
            .month = strSplt(1)
            .ndate = strSplt(2)
            .parts = strSplt(3)
            .kum = chrStr(strSplt(4))
            .danga = chrStr(strSplt(5))
            .km = chrStr(strSplt(6))
            .oil = chrStr(strSplt(7))
            .rkm = chrStr(strSplt(8))
            .km1L = chrStr(strSplt(9))
            .store = strSplt(10)
            .etc = strSplt(11)
        End With
    End Sub
    Public Function chrStr(ByRef strCh As String)
        '분리한 데이터에서 필요 없는 것 지우기
        Dim s As String
        s = Replace(strCh, ChrW(34), "")
        s = Replace(s, "?", "")
        s = Replace(s, " ", "")

        Return s
    End Function
    Public Sub sltGrid()
        '메인 화면의 DataGridView에 넣기_메인 소스 데이터
        Dim i As Integer = -1

        'Main.dtGrip_List.RowCount = iCounter + 1
        Main.dtGrip_List.RowCount = UBound(cartotal1)
        For i = 0 To iCounter - 1
            With Main.dtGrip_List
                .Rows(i).Cells(0).Value = cartotal1(i).month
                .Rows(i).Cells(1).Value = cartotal1(i).ndate
                .Rows(i).Cells(2).Value = cartotal1(i).parts
                .Rows(i).Cells(3).Value = rtnComma(cartotal1(i).kum)
                .Rows(i).Cells(4).Value = rtnComma(cartotal1(i).danga)
                .Rows(i).Cells(5).Value = rtnComma(cartotal1(i).km)
                .Rows(i).Cells(6).Value = rtnComma(cartotal1(i).oil)
                .Rows(i).Cells(7).Value = rtnComma(cartotal1(i).rkm)
                .Rows(i).Cells(8).Value = rtnComma(cartotal1(i).km1L)
                .Rows(i).Cells(9).Value = cartotal1(i).store
                .Rows(i).Cells(10).Value = cartotal1(i).etc

            End With

            Main.dtGrip_List.CurrentCell = Main.dtGrip_List.Rows(i).Cells(0)
        Next

    End Sub

    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 주유소 리스트 Reading Store and Address
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    Public Sub readStoreFile()
        '메인 데이터 소스 읽어 들이기
        'FileOpen(1, dataFolder & "\StoreList.txt", OpenMode.Input, OpenAccess.Read)
        'str = LineInput(1)
        'Do Until EOF(1)

        '    iStoreCounter += 1
        '    str = LineInput(1)
        '    sltDataStore(str)

        'Loop

        Using sr1 As StreamReader = New StreamReader(dataFolder & "StoreList.txt", System.Text.Encoding.UTF8)
            str = sr1.ReadLine

            Do While str <> Nothing
                iStoreCounter += 1
                str = sr1.ReadLine

                If str = Nothing Then
                    'MsgBox("없다")
                    iStoreCounter -= 1
                    GoTo RdEnd
                End If

                If Len(str) > 30 Then
                    sltDataStore(str)
                Else
                    iStoreCounter -= 1
                End If
            Loop
RdEnd:
        End Using
        sltGridStore()

        FileClose()
    End Sub
    Public Sub sltDataStore(ByRef strDt As String)
        '파일 데이터 읽어 분리하는 루틴
        ReDim Preserve subStore1(iStoreCounter)
        strSplt = Split(strDt, spc)

        With subStore1(iStoreCounter - 1)
            .store = strSplt(0)
            .etc = strSplt(1)
        End With
    End Sub
   
    Public Sub sltGridStore()
        '메인 화면의 DataGridView에 넣기_메인 소스 데이터
        Dim i As Integer = -1

        subStoreList.subGrid.RowCount = iStoreCounter + 1
        For i = 0 To iStoreCounter - 1
            With subStoreList.subGrid
                .Rows(i).Cells(0).Value = subStore1(i).store
                .Rows(i).Cells(1).Value = subStore1(i).etc

            End With
        Next
    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 부품 수리 내역 리스트
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    Public Parts1() As Parts

    Public Sub readParts()
        '메인 데이터 소스 읽어 들이기
        'FileOpen(1, dataFolder & "SpartsStore.txt", OpenMode.Input, OpenAccess.Read)
        'str = LineInput(1)
        'Do Until EOF(1)

        '    iParts += 1
        '    str = LineInput(1)
        '    sltDataParts(str)
        'Loop


        Using sr As StreamReader = New StreamReader(dataFolder & "SpartsStore.txt", System.Text.Encoding.UTF8)
            str = sr.ReadLine
            Do While str <> Nothing
                iParts += 1
                str = sr.ReadLine

                If str = Nothing Then
                    'MsgBox("없다")
                    iParts -= 1
                    GoTo RdEnd
                End If

                If Len(str) > 30 Then
                    sltDataParts(str)
                Else
                    iParts -= 1
                End If

            Loop

RdEnd:
        End Using
        sltRepairList()
        FileClose()
    End Sub
    Public Sub sltDataParts(ByRef strDt As String)
        '파일 데이터 읽어 분리하는 루틴
        ReDim Preserve Parts1(iParts)
        strSplt = Split(strDt, spc)

        With Parts1(iParts - 1)
            .year = chrStr(strSplt(0))
            .month = chrStr(strSplt(1))
            .pdate = chrStr(strSplt(2))
            .part = strSplt(3)
            .kum = chrStr(strSplt(4))
            .km = chrStr(strSplt(5))
            .rkm = chrStr(strSplt(6))
            .store = Replace(strSplt(7), ChrW(34), "")
            .etc = Replace(strSplt(8), ChrW(34), "")
        End With
    End Sub
    Public Sub sltRepairList()
        '메인 화면의 DataGridView에 넣기_메인 소스 데이터
        Dim i As Integer = -1

        Main.gbGridStore.RowCount = iParts + 1
        For i = 0 To iParts - 1

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

    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 부품 수리점 주소 리스트
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    Public PartAddressList() As subStore

    Public Sub readPartsAddress()
        '메인 데이터 소스 읽어 들이기
        'FileOpen(1, dataFolder & "spartsStorelist.txt", OpenMode.Input, OpenAccess.Read)
        'str = LineInput(1)
        'Do Until EOF(1)

        '    iPartsStore += 1
        '    str = LineInput(1)
        '    sltDataPartsAddress(str)
        'Loop

        Using sr As StreamReader = New StreamReader(dataFolder & "spartsStorelist.txt", System.Text.Encoding.UTF8)
            str = sr.ReadLine
            Do While str <> Nothing
                iPartsStore += 1
                str = sr.ReadLine

                If str = Nothing Then
                    'MsgBox("없다")
                    iPartsStore -= 1
                    GoTo RdEnd
                End If

                If Len(str) > 15 Then
                    sltDataPartsAddress(str)
                Else
                    iPartsStore -= 1
                End If
            Loop

RdEnd:
        End Using

        sltPartsGrid()

        FileClose()
    End Sub
    Public Sub sltDataPartsAddress(ByRef strDt As String)
        '파일 데이터 읽어 분리하는 루틴
        ReDim Preserve PartAddressList(iPartsStore)
        strSplt = Split(strDt, spc)

        With PartAddressList(iPartsStore - 1)
            .store = chrStr(strSplt(0))
            .etc = chrStr(strSplt(1))

        End With
    End Sub
    Public Sub sltPartsGrid()
        '수리점 과 주소를 옮기기 (form : subPartsList)
        Dim i As Integer = -1

        subPartsList.subGrid.RowCount = iPartsStore + 1
        For i = 0 To iPartsStore - 1
            With subPartsList.subGrid.Rows(i)
                .Cells(0).Value = PartAddressList(i).store
                .Cells(1).Value = PartAddressList(i).etc
            End With
        Next
    End Sub

    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 수리 부품 총괄  리스트
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    Public repairParts1() As repairParts
    Public cntParts As Integer = 0
    Public Sub readRepairPartsList()
        '메인 데이터 소스 읽어 들이기
        'FileOpen(1, dataFolder & "Spartslist.txt", OpenMode.Input, OpenAccess.Read)
        'str = LineInput(1)
        'Do Until EOF(1)

        '    cntParts += 1
        '    str = LineInput(1)
        '    sltRepairParts(str)
        'Loop

        Using sr As StreamReader = New StreamReader(dataFolder & "Spartslist.txt", System.Text.Encoding.UTF8)
            str = sr.ReadLine
            Do While str <> Nothing
                cntParts += 1
                str = sr.ReadLine

                If str = Nothing Then
                    'MsgBox("없다")
                    cntParts -= 1
                    GoTo RdEnd
                End If

                If Len(str) > 5 Then
                    sltRepairParts(str)
                Else
                    cntParts -= 1
                End If
            Loop

RdEnd:
        End Using
        sltRepairGrid()
        FileClose()
    End Sub
    Public Sub sltRepairParts(ByRef strDt As String)
        '파일 데이터 읽어 분리하는 루틴
        ReDim Preserve repairParts1(cntParts)
        strSplt = Split(strDt, spc)
        'strSplt = Split(strDt, ",")

        With repairParts1(cntParts - 1)
            .name = strSplt(0)
            .chgKm = chrStr(strSplt(1))
            .move = chrStr(strSplt(2))
            .pdate = chrStr(strSplt(3))

            .name1 = strSplt(4)
            .chgKm1 = chrStr(strSplt(5))
            .move1 = chrStr(strSplt(6))
            .pdate1 = chrStr(strSplt(7))

            .name2 = strSplt(8)
            .chgKm2 = chrStr(strSplt(9))
            .move2 = chrStr(strSplt(10))
            .pdate2 = chrStr(strSplt(11))
        End With
    End Sub
    Public Sub sltRepairGrid()
        '메인 화면의 DataGridView에 넣기_메인 소스 데이터
        Dim i As Integer = -1

        'Main.dtGrip_List.RowCount = 3
        For i = 0 To cntParts - 1

            With Main.gbSub_dtGrip
                .Rows(i).Cells(0).Value = repairParts1(i).name
                .Rows(i).Cells(1).Value = rtnComma(repairParts1(i).chgKm)
                .Rows(i).Cells(2).Value = rtnComma(repairParts1(i).move)
                .Rows(i).Cells(3).Value = repairParts1(i).pdate

                .Rows(i).Cells(4).Value = repairParts1(i).name1
                .Rows(i).Cells(5).Value = rtnComma(repairParts1(i).chgKm1)
                .Rows(i).Cells(6).Value = rtnComma(repairParts1(i).move1)
                .Rows(i).Cells(7).Value = repairParts1(i).pdate1

                .Rows(i).Cells(8).Value = repairParts1(i).name2
                .Rows(i).Cells(9).Value = rtnComma(repairParts1(i).chgKm2)
                .Rows(i).Cells(10).Value = rtnComma(repairParts1(i).move2)
                .Rows(i).Cells(11).Value = repairParts1(i).pdate2
            End With
        Next

    End Sub

    '/////////////////////////////////////////////////////////////////////////////////////////////////
    '///// 총 결과  리스트
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    Public Total1 As Total

    Public Sub readTotalCumsumer()
        '메인 데이터 소스 읽어 들이기
        'FileOpen(1, dataFolder & "Stotalcar.txt", OpenMode.Input, OpenAccess.Read)
        'str = LineInput(1)
        'Do Until EOF(1)

        '    cntParts += 1
        '    str = LineInput(1)
        '    sltTotal(str)
        'Loop

        Using sr As StreamReader = New StreamReader(dataFolder & "Stotalcar.txt", System.Text.Encoding.UTF8)
            str = sr.ReadLine
            str = sr.ReadLine

                sltTotal(str)
         
        End Using
        sltTotalGrid()
        FileClose()
    End Sub
    Public Sub sltTotal(ByRef strDt As String)
        '파일 데이터 읽어 분리하는 루틴

        strSplt = Split(strDt, spc)

        With Total1
            .Kum = strSplt(0)
            .ttlLiter = chrStr(strSplt(1))
            .ttlKm = chrStr(strSplt(2))
            .avgKm = chrStr(strSplt(3))
            .wonofLiter = strSplt(4)
            
        End With
    End Sub
    Public Sub sltTotalGrid()
        '메인 화면의 DataGridView에 넣기_메인 소스 데이터
        Dim i As Integer = -1

        'Main.dtGrip_List.RowCount = 3

        With Total1
            Main.gbtTx_MTotal.Text = rtnComma(.Kum)
            Main.gbtTx_TLiter.Text = rtnComma(.ttlLiter)
            Main.gbtTx_TTLKm.Text = rtnComma(.ttlKm)
            Main.gbtTx_aveOil.Text = rtnComma(.avgKm)
            Main.gbtTx_won1L.Text = rtnComma(.wonofLiter)
           
        End With
    End Sub
End Module
