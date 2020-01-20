'//////////////////////////////////////////////////////////////////////////////////////////////////////
'///////// 메인 데이터 화면
'//////////////////////////////////////////////////////////////////////////////////////////////////////
Imports System.IO
Imports System
Imports System.Windows.Forms

Public Class Main

    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '로그램 종료로 가는 길 목

        '파일 저장
        svLstFile()  '주유 내역
        svStoreFile()  '주유소 리스트
        svParts() '부품 수리 내역
        svRepairPartsList() ' 수리 부품 내역 총괄 리스트
        svTotalCumsumer() '총 결과 리스트
        svRepairFile() '정비소 리스트 저장
    End Sub

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.gbSub_dtGrip.RowCount = 3
        Me.dtGrip_List.RowCount = 20

        readFile() '파일 읽기
        readStoreFile() ' 주유소 주소 읽기
        readParts() ' 부품 수리 내용 리스트
        readPartsAddress() '수리점 리스트
        readRepairPartsList() '부붐 수리 내용 리스트

        readTotalCumsumer() '총합 결과

        Me.gbtCb_Month.SelectedIndex = rtnMonth() - 1
        'Me.gbtCb_Parts.SelectedIndex = 0
    End Sub

    Private Sub gbtCb_Parts_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbtCb_Parts.SelectedIndexChanged
        '경유와 기타 부품들의 선택
        Dim i As Integer = -1
       
        i = Me.gbtCb_Parts.SelectedIndex
        If i = 0 Then
            Me.dtGrip_List.Visible = True
            Me.gbGridStore.Visible = False
            subStoreList.ShowDialog()
        Else

            Me.gbGridStore.Visible = True
            Me.dtGrip_List.Visible = False
            subPartsList.ShowDialog()
        End If
    End Sub

   
    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        '데이터 옮기기
        Dim i As Integer = -1

        My.Computer.Keyboard.SendKeys(Keys.HanguelMode, True)

        i = Me.gbtCb_Parts.SelectedIndex
        'movePartsLen() '부품 교환후 이동 거리
        If i = 0 Then
            moveData() '입력한 데이터를 옮기기  Main
            moveOil() '주유 리스트 옮기기  Main

            movePartsLen() ' '부품 교환후 이동 거리
        Else
            movePartsLen() ' '부품 교환후 이동 거리
            movePart() '부품 교환 리스트 옮기기  Main
            moveParts() '입력한 데이터를 교환 부품 리스트 옮기기  Main
        End If
        'movePartsLen() ' '부품 교환후 이동 거리
        moveTotal() '입력한 데이터를 총괄 결과 데이터로 옮기기  Main
    End Sub

    Private Sub gbtTx_TTLKm_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbtTx_TTLKm.TextChanged
        '총 이동 거리
        '평균 리터당 단가 계산
        Dim i As Integer = -1
        Dim inStr As String

        inStr = Me.gbtTx_TTLKm.Text
        If inStr.Length > 0 Then
            i = Me.gbtCb_Parts.SelectedIndex
            If i = 0 Then Me.gbtTx_aveOil.Text = Format(CDbl(Me.gbtTx_TTLKm.Text) / CDbl(Me.gbtTx_TLiter.Text), "###,###.#0")
        End If

       
    End Sub

    Private Sub gbtTx_TLiter_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbtTx_TLiter.TextChanged
        '총 주유량
        '평균 연비 계산
        Dim i As Integer = -1
        Dim inStr As String

        inStr = Me.gbtTx_TLiter.Text

        If inStr.Length > 0 Then
            i = Me.gbtCb_Parts.SelectedIndex

            If i = 0 Then Me.gbtTx_won1L.Text = Format(Me.gbtTx_MTotal.Text / CDbl(Me.gbtTx_TLiter.Text), "###,###.#0")
        End If
    End Sub

    Private Sub gbtTx_Danga_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbtTx_Danga.Leave
        'Me.gbtTx_Oil.Text = Format(CDbl(Me.gbtTx_Oil.Text), "###.#0")
        '주유 단가 입력
        '주유량 계산
        Dim i As Integer = -1
        Dim inStr As String
        Dim inStr1 As String = rtnNotComma(Me.gbtTx_Kum.Text)
        Dim inStr2 As String = rtnNotComma(Me.gbtTx_Danga.Text)

        Try

            inStr = Me.gbtTx_Danga.Text

            If inStr.Length > 0 Then
                i = Me.gbtCb_Parts.SelectedIndex
                If i = -1 Then GoTo err_openParts
                If i = 0 Then Me.gbtTx_Oil.Text = Format(CDbl(inStr1) / CDbl(inStr2), "###.#0")
            Else
                MsgBox("빈칸이다")
            End If

            Me.gbtTx_Kum.Text = rtnComma(Me.gbtTx_Kum.Text)
            Me.gbtTx_Danga.Text = rtnComma(Me.gbtTx_Danga.Text)
            Exit Sub
err_OpenParts:
            MsgBox("부품을 선택하여 주십시오.")
            'Me.gbtCb_Parts.Enabled = True
        Catch ex As Exception

        End Try
    End Sub

    'Private Sub gbtTx_Danga_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbtTx_Danga.TextChanged
    '    ''주유 단가 입력
    '    ''주유량 계산
    '    'Dim i As Integer = -1
    '    'Dim inStr As String
    '    'Dim inStr1 As String = rtnNotComma(Me.gbtTx_Kum.Text)
    '    'Dim inStr2 As String = rtnNotComma(Me.gbtTx_Danga.Text)

    '    'inStr = Me.gbtTx_Danga.Text

    '    'If inStr.Length > 0 Then
    '    '    i = Me.gbtCb_Parts.SelectedIndex
    '    '    If i = 0 Then Me.gbtTx_Oil.Text = Format(CDbl(inStr1) / CDbl(inStr2), "###.#0")
    '    'Else
    '    '    MsgBox("빈칸이다")
    '    'End If

    '    'Me.gbtTx_Oil.Text = Format(CDbl(Me.gbtTx_Oil.Text), "###.#0")
    'End Sub

    Private Sub gbtTx_Km_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbtTx_Km.Leave
        'Me.gbtTx_TKm.Text = Format(CDbl(Me.gbtTx_TKm.Text), "#,###,###.##")
        '이동 거리
        '앞 주유한 거리에서 이번 거리까지의 실제 이돈 거리
        Try

            Dim i As Integer = -1
            Dim inStr As String
            Dim inStr1 As String = rtnNotComma(Me.gbtTx_Km.Text)
            Dim inStr2 As String = rtnNotComma(cartotal1(iCounter - 1).km)

            inStr = Me.gbtTx_Km.Text

            'If inStr.Length > 0 Then
            i = Me.gbtCb_Parts.SelectedIndex
            If i = -1 Then GoTo err_openParts
            If i = 0 Then Me.gbtTx_TKm.Text = Format(CDbl(inStr1) - CDbl(inStr2), "#,###,###.#0")
            
            Me.gbtTx_Km.Text = rtnComma(Me.gbtTx_Km.Text)
            Exit Sub
err_OpenParts:
            MsgBox("부품을 선택하여 주십시오.")
            'Me.gbtCb_Parts.Enabled = True
        Catch ex As Exception


        End Try

    End Sub

    'Private Sub gbtTx_Km_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbtTx_Km.TextChanged
    '    ''이동 거리
    '    ''앞 주유한 거리에서 이번 거리까지의 실제 이돈 거리
    '    'Dim i As Integer = -1
    '    'Dim inStr As String

    '    'inStr = Me.gbtTx_Km.Text

    '    ''If inStr.Length > 0 Then
    '    'i = Me.gbtCb_Parts.SelectedIndex
    '    'If i = 0 Then
    '    '    Me.gbtTx_TKm.Text = Format(Me.gbtTx_Km.Text - cartotal1(iCounter).km, "#,###,###.#0")
    '    'Else
    '    '    MsgBox("빈칸이다")
    '    'End If


    '    'Me.gbtTx_TTLKm.Text = Me.gbtTx_Km.Text
    '    'End If
    'End Sub

    Private Sub gbtTx_KmL_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbtTx_KmL.TextChanged
        'Liter당 거리 (km)
        Dim i As Integer
        i = Me.gbtCb_Parts.SelectedIndex
        If i = 0 Then Me.gbtTx_KmL.Text = Format(CDbl(Me.gbtTx_KmL.Text), "###,###.##")
    End Sub

    Private Sub gbtTx_TKm_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbtTx_TKm.Leave
        '리터당 거리는 오일일 떄만 입력
        'Dim i As Integer
        'i = Me.gbtCb_Parts.SelectedIndex
        'If i = 0 Then
        Me.gbtTx_KmL.Text = Format(CDbl(Me.gbtTx_KmL.Text), "##,###.#0")
    End Sub

    Private Sub gbtTx_TKm_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbtTx_TKm.TextChanged
        '실제 이동 거리
        ' 'Liter당 거리 (km) 계산
        Dim i As Integer = -1
        Dim inStr As String

        inStr = Me.gbtTx_TKm.Text

        If inStr.Length > 0 Then
            i = Me.gbtCb_Parts.SelectedIndex
            If i = 0 Then Me.gbtTx_KmL.Text = Format(CDbl(Me.gbtTx_TKm.Text) / CDbl(cartotal1(iCounter - 1).oil), "##,###.#0")
        End If
    End Sub

    

    Private Sub mnuEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEnd.Click
        ' 프로그램 종료
        Me.IconEnd.PerformClick()

    End Sub

    Private Sub mnuXls_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuXls.Click
        '엑셀 저장
        Me.IconExcel.PerformClick()
    End Sub

    Private Sub mnuTxt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTxt.Click
        '텍스트 저장
        Me.IconTxt.PerformClick()
    End Sub

    Private Sub mnuOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpen.Click
        '파일 읽기
        Me.IconOpen.PerformClick()
    End Sub

    Private Sub mnuNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNew.Click
        '새로이 
        Me.IconNew.PerformClick()
    End Sub

   
    Private Sub mnuVer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuVer.Click
        '프로그램 버젼
        ver.ShowDialog()
    End Sub

    Private Sub mnuDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDate.Click
        ' 자동차 구입 시기
        carDate.ShowDialog()
    End Sub

    Private Sub mnuCarNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCarNo.Click
        '자동차 등록 번호
        carNumber.ShowDialog()
    End Sub


    Private Sub mnuCars_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCars.Click
        '자동차 종류
        carName.ShowDialog()
    End Sub
    Private Sub mnuOil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOil.Click
        '주유 입력 - 경유 구자입 입력
        Me.IconOil.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 0
    End Sub
    Private Sub mnuEnginOil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEnginOil.Click
        '엔진오일 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 1
    End Sub
   
    Private Sub mnuBrkPad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBrkPad.Click
        '브레이크 패드 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 2
    End Sub
    Private Sub mnuEngin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEngin.Click
        '부동액 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 3
    End Sub
    Private Sub mnuTireBelt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTireBelt.Click
        '타이어 밸트 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 4
    End Sub
    Private Sub mnuAircon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAircon.Click
        '에어컨 가스 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 5
    End Sub
    Private Sub mnuIceBelt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIceBelt.Click
        '냉각수 밸트 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 6
    End Sub
    Private Sub mnuAirFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAirFilter.Click
        '에어필터 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 7
    End Sub
    Private Sub mnuBattery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBattery.Click
        '밧데리 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 8
    End Sub
    Private Sub mnuTire_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTire.Click
        '타이어 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 9
    End Sub
    Private Sub mnuMission_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMission.Click
        '미션오일 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 10
    End Sub
    Private Sub mnuShobar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShobar.Click
        '쇼바 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 11
    End Sub
    Private Sub mnuGasGet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGasGet.Click
        '가스켓 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 12
    End Sub

    Private Sub mnuJoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuJoint.Click
        '등속 조인트 입력
        Me.IconParts.PerformClick()
        Me.gbtCb_Parts.SelectedIndex = 13
    End Sub

   

    Private Sub IconParts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IconParts.Click
        '자동차 부품 수리 내용
        MsgBox("자동차 부품 수리 내역")
    End Sub

    Private Sub IconOil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IconOil.Click
        '자동차 주유 경유 구입
        MsgBox("자동차 주유 경유 구입")
    End Sub

    Private Sub IconExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IconExcel.Click
        '엑셀 저장
        MsgBox("엑셀 저장")
    End Sub

    Private Sub IconTxt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IconTxt.Click
        '텍스트 저장
        MsgBox("텍스트 저장")
    End Sub

    Private Sub IconOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IconOpen.Click
        '파일 읽기
        MsgBox("파일 열기")
    End Sub

    Private Sub IconNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IconNew.Click
        '작업 새로이
        MsgBox("작업 새로이")
    End Sub

    Private Sub IconEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IconEnd.Click
        '프로그램 종료

        '파일 저장
        svLstFile()  '주유 내역
        svStoreFile()  '주유소 리스트
        svParts() '부품 수리 내역
        svRepairPartsList() ' 수리 부품 내역 총괄 리스트
        svTotalCumsumer() '총 결과 리스트
        svRepairFile() '정비소 리스트 저장

        End
    End Sub

   
    Private Sub gbtTx_MTotal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gbtTx_MTotal.TextChanged
        'Me.gbtTx_MTotal.Text = Format(Me.gbtTx_MTotal.Text, "###,###,###,##0")
    End Sub

    Private Sub gbtTx_aveOil_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbtTx_aveOil.Leave
        Me.gbtTx_aveOil.Text = Format(CDbl(Me.gbtTx_aveOil.Text), "###,###.#0")
    End Sub


    Private Sub gbtTx_won1L_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbtTx_won1L.Leave
        Me.gbtTx_won1L.Text = Format(CDbl(Me.gbtTx_won1L.Text), "###,###.#0")
    End Sub

    Private Sub gbtTx_Kum_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbtTx_Kum.Leave
        Me.gbtTx_Kum.Text = Format(CDbl(Me.gbtTx_Kum.Text), "###,###.###")
    End Sub


    Private Sub gbtTx_Oil_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles gbtTx_Oil.Leave
        '오일일 때만 값이 입력
        Dim I As Integer
        I = Me.gbtCb_Parts.SelectedIndex
        If I = 0 Then Me.gbtTx_Oil.Text = Format(CDbl(Me.gbtTx_Oil.Text), "###,###.###")
    End Sub

End Class
