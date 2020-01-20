'//////////////////////////////////////////////////////////////////////////////////////////////////////
'///////// 기본 선언 변수
'//////////////////////////////////////////////////////////////////////////////////////////////////////



Module standard

    Public rCounter As Integer = 0 '현재 카운터
    Public iCounter As Integer = 0 '임시 카운터
    Public iStoreCounter As Integer = 0 '임시 주유소와 주소 카운터
    Public iParts As Integer = 0 '부품 수리 리스트
    Public iPartsStore As Integer = 0 '부품점 , 주소 리스트

    Public Structure CarTotal '메인 데이터 소스
        Public year As String
        Public month As String
        Public ndate As String
        Public parts As String
        Public kum As String
        Public danga As String
        Public km As String
        Public oil As String
        Public rkm As String
        Public km1L As String
        Public store As String
        Public etc As String
    End Structure

    Public Structure subStore '주유소, 주소 리스트
        Public store As String
        Public etc As String
    End Structure

    Public Structure Parts '부품 수리 리스트 
        Public year As String
        Public month As String
        Public pdate As String
        Public part As String
        Public kum As String
        Public km As String
        Public rkm As String
        Public store As String
        Public etc As String
    End Structure

    Public Structure repairParts '부품 수리 결과 내용
        Public name As String
        Public chgKm As String
        Public move As String
        Public pdate As String

        Public name1 As String
        Public chgKm1 As String
        Public move1 As String
        Public pdate1 As String

        Public name2 As String
        Public chgKm2 As String
        Public move2 As String
        Public pdate2 As String
    End Structure

    Public Structure Total '총합 내용
        Public Kum As String
        Public ttlLiter As String
        Public ttlKm As String
        Public avgKm As String
        Public wonofLiter As String
    End Structure

    Public Function rtnMonth() '금일 월을 선택
        Dim iMonth As String
        Dim iMnth() As String

        iMonth = Format(Now, "YYYY-MM-DD")
        iMnth = Split(iMonth, "-")

        Return iMnth(1)
    End Function

    Public Function rtnComma(ByRef rtnStr As String) '숫자에 comma (,)를 넣은 것


        Select Case rtnStr
            Case "-"
                rtnStr = "0"
            Case "#REF!"
                rtnStr = "0"

            Case ""
                rtnStr = "0"
            Case Else
                rtnStr = Format(CDbl(rtnStr), "###,###.###")
        End Select
        'If "-" <> rtnStr Then
        'ElseIf "#REF!" <> rtnStr Then
        '    rtnStr = Format(CDbl(rtnStr), "###,###.###")

        'Else
        '    rtnStr = "0"

        'End If

        Return rtnStr
    End Function
    Public Function rtnNotComma(ByRef rtnStr As String) '숫자에 comma (,)를 없는 것

        Try

            rtnStr = Replace(rtnStr, ",", "")
            Return rtnStr

        Catch ex As Exception
            Return "0"
        End Try
    End Function

  
End Module
