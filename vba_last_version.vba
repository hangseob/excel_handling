Sub CreateTimeFwdChart_Last()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim targetTbl As ListObject
    Dim colStart As Integer, colEnd As Integer
    Dim colFwdStart As Integer, colFwdEnd As Integer
    Dim i As Long, targetRow As Long
    Dim newTbl As ListObject
    Dim chartObj As ChartObject
    Dim foundTable As Boolean
    
    ' 1. 현재 활성화된 시트(ActiveSheet) 설정
    Set wsSource = ActiveSheet
    foundTable = False
    
    ' MarketTable로 시작하는 표 찾기
    For Each targetTbl In wsSource.ListObjects
        If Left(targetTbl.Name, 11) = "MarketTable" Then
            foundTable = True
            Exit For
        End If
    Next targetTbl
    
    If Not foundTable Then
        MsgBox "현재 시트에서 'MarketTable'로 시작하는 테이블을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 2. 컬럼명으로 인덱스 찾기 (유연한 대응)
    On Error Resume Next
    colStart = targetTbl.ListColumns("구간 시작").Index
    colEnd = targetTbl.ListColumns("구간 끝").Index
    colFwdStart = targetTbl.ListColumns("fwd 시작").Index
    colFwdEnd = targetTbl.ListColumns("fwd 끝").Index
    On Error GoTo 0
    
    If colStart = 0 Or colEnd = 0 Or colFwdStart = 0 Or colFwdEnd = 0 Then
        MsgBox "테이블 내에 필요한 컬럼명이 없습니다." & vbCrLf & _
               "(필요: 구간 시작, 구간 끝, fwd 시작, fwd 끝)", vbExclamation
        Exit Sub
    End If
    
    ' 3. 새로운 워크시트를 현재 시트(wsSource) 바로 오른쪽에 추가
    Set wsNew = Sheets.Add(After:=wsSource)
    wsNew.Name = "Last_Chart_" & Format(Now, "HHMMSS")
    
    ' 4. 새로운 표 헤더 추가
    wsNew.Cells(1, 1).Value = "time"
    wsNew.Cells(1, 2).Value = "fwd"
    wsNew.Columns("A:A").NumberFormat = "yyyy-mm-dd" ' 데이터 열 자체도 날짜 형식으로 지정
    
    targetRow = 2
    
    ' 5. 테이블 데이터를 순회하며 세로로 배치 (DataBodyRange 사용)
    For i = 1 To targetTbl.ListRows.Count
        ' 구간 시작 & fwd 시작
        wsNew.Cells(targetRow, 1).Value = targetTbl.DataBodyRange(i, colStart).Value
        wsNew.Cells(targetRow, 2).Value = targetTbl.DataBodyRange(i, colFwdStart).Value
        targetRow = targetRow + 1
        
        ' 구간 끝 & fwd 끝
        wsNew.Cells(targetRow, 1).Value = targetTbl.DataBodyRange(i, colEnd).Value
        wsNew.Cells(targetRow, 2).Value = targetTbl.DataBodyRange(i, colFwdEnd).Value
        targetRow = targetRow + 1
    Next i
    
    ' 6. 결과 데이터를 표(Table)로 변환
    Set newTbl = wsNew.ListObjects.Add(xlSrcRange, wsNew.Range("A1:B" & targetRow - 1), , xlYes)
    newTbl.Name = "ResultTable_" & Format(Now, "HHMMSS")
    newTbl.TableStyle = "TableStyleMedium2"
    
    ' 7. 분산형 차트 생성
    Set chartObj = wsNew.ChartObjects.Add(Left:=wsNew.Cells(1, 4).Left, Width:=500, Top:=wsNew.Cells(1, 4).Top, Height:=350)
    
    With chartObj.Chart
        .ChartType = xlXYScatterLines
        .SetSourceData Source:=newTbl.Range
        .HasTitle = True
        .ChartTitle.Text = "Time vs FWD (" & targetTbl.Name & ")"
        
        ' 축 제목 및 상세 설정
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Time"
            .TickLabels.NumberFormat = "yyyy-mm-dd" ' 날짜 형식 적용
            .TickLabels.Orientation = 90            ' 90도 회전
            .HasMajorGridlines = True               ' 주 그리드선
            .HasMinorGridlines = True               ' 보조 그리드선
        End With
        
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "FWD"
            .HasMajorGridlines = True
        End With
        
        .HasLegend = False
    End With
    
    wsNew.Columns("A:B").AutoFit
    MsgBox "ActiveSheet의 바로 오른쪽에 시트를 생성하고 차트를 만들었습니다.", vbInformation
End Sub
