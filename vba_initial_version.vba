Sub CreateTimeFwdChart_Initial()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim i As Long, targetRow As Long
    Dim tbl As ListObject
    Dim chartObj As ChartObject
    
    ' 현재 데이터가 있는 시트 설정
    Set wsSource = ActiveSheet
    
    ' 원본 데이터의 마지막 행 찾기 (A열 기준)
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' 데이터가 없는 경우 종료
    If lastRow < 2 Then
        MsgBox "데이터가 없습니다. 1행은 제목, 2행부터 데이터가 있어야 합니다.", vbExclamation
        Exit Sub
    End If
    
    ' 새로운 워크시트를 맨 마지막에 추가
    Set wsNew = Sheets.Add(After:=Sheets(Sheets.Count))
    wsNew.Name = "Initial_Data_" & Format(Now, "HHMMSS")
    
    ' 새로운 표 헤더 추가
    wsNew.Cells(1, 1).Value = "time"
    wsNew.Cells(1, 2).Value = "fwd"
    
    targetRow = 2
    
    ' 원본 데이터를 순회하며 세로로 배치 (A, B, C, D열 고정)
    For i = 2 To lastRow
        ' 1. 구간 시작 & fwd 시작 배치
        wsNew.Cells(targetRow, 1).Value = wsSource.Cells(i, 1).Value
        wsNew.Cells(targetRow, 2).Value = wsSource.Cells(i, 3).Value
        targetRow = targetRow + 1
        
        ' 2. 구간 끝 & fwd 끝 배치
        wsNew.Cells(targetRow, 1).Value = wsSource.Cells(i, 2).Value
        wsNew.Cells(targetRow, 2).Value = wsSource.Cells(i, 4).Value
        targetRow = targetRow + 1
    Next i
    
    ' 데이터를 표(Table) 형식으로 변환
    Set tbl = wsNew.ListObjects.Add(xlSrcRange, wsNew.Range("A1:B" & targetRow - 1), , xlYes)
    tbl.Name = "InitialTable_" & Format(Now, "HHMMSS")
    tbl.TableStyle = "TableStyleMedium2"
    
    ' 분산형 차트 생성
    Set chartObj = wsNew.ChartObjects.Add(Left:=wsNew.Cells(1, 4).Left, Width:=500, Top:=wsNew.Cells(1, 4).Top, Height:=350)
    
    With chartObj.Chart
        .ChartType = xlXYScatterLines
        .SetSourceData Source:=tbl.Range
        .HasTitle = True
        .ChartTitle.Text = "Time vs FWD (Initial Version)"
        
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Time"
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "FWD"
        End With
        
        .HasLegend = False
    End With
    
    wsNew.Columns("A:B").AutoFit
End Sub
