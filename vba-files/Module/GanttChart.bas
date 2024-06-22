Attribute VB_Name = "GanttChart"

Public Sub 実績線を描画する()

    Application.ScreenUpdating = False

    Dim startDate, endDate As Date                      ' 開始日, 終了日
    Dim percentage As Double                            ' 進捗率
    Dim square As Shape                                 ' 実績線長方形
    Dim leftPoint, topPoint, height, width As Double    ' 実績線サイズ
    Dim statusColumn, startDateColumn, endDateColumn, percentageColumn As Integer
    Dim status As String
    Dim i As Integer
    Dim startDateCell As Range, endDateCell As Range
    
    ' 実績線を削除する
    Worksheets("WBS").Activate
    For Each square In ActiveSheet.Shapes
        If square.Name = "実績線" Then
            square.Delete
        End If
    Next
    
    statusColumn = Range("_status").Column              ' 状況 列
    startDateColumn = Range("_startDate").Column        ' 開始日 列
    endDateColumn = Range("_endDate").Column            ' 終了日 列
    percentageColumn = Range("_progress").Column        ' 進捗率 列
    
    For i = Range("_taskStart").Row To Range("_taskStart").End(xlDown).Row

        ' 非表示行の場合スキップ
        If Rows(i).Hidden Then
            GoTo Continue
        End If

        ' ステータスが対応完了、対応中、保留でない場合
        status = Cells(i, statusColumn).Value
        If status <> "対応完了" And status <> "対応中" And status <> "保留" Then
            GoTo Continue
        End If
        
        ' 開始日、終了日、進捗率を取得
        startDate = Cells(i, startDateColumn).Value
        endDate = Cells(i, endDateColumn).Value
        percentage = Cells(i, percentageColumn).Value
        
        ' カレンダー行から開始日セル、終了日セルを取得
        Set startDateCell = Rows(Range("_calendar").Row).Find(What:=DateValue(startDate), LookIn:=xlFormulas)
        Set endDateCell = Rows(Range("_calendar").Row).Find(What:=DateValue(endDate), LookIn:=xlFormulas)
        If startDateCell Is Nothing Or endDateCell Is Nothing Then
            GoTo Continue
        End If
        
        Dim startColumn As Integer, endColumn As Integer
        startColumn = startDateCell.Column
        endColumn = endDateCell.Column
        
        ' 実績線図形のサイズ・位置設定
        leftPoint = Cells(i, startColumn).Left                                                  ' 左位置
        topPoint = Cells(i, startColumn).Top                                                    ' 上位置
        height = Cells(i, startColumn).height - 10                                              ' 高さ
        width = (Cells(i, endColumn + 1).Left - Cells(i, startColumn).Left) * percentage        ' 幅
        
        ' 実績線を作成
        Set square = ActiveSheet.Shapes.AddShape(msoShapeRectangle, leftPoint, topPoint, width, height)
        square.Fill.ForeColor.SchemeColor = 32              ' 色：青
        square.Name = "実績線"                              ' 名称
        square.Fill.Transparency = 0.4                      ' 透明度：40%
        
Continue:
    Next i
    
    Application.ScreenUpdating = True

End Sub


Sub 稲妻線を描画する()

    Application.ScreenUpdating = False

    Dim line As Shape
    
    Worksheets("WBS").Activate

    ' 既存の稲妻線を削除
    For Each line In ActiveSheet.Shapes
        If line.Name = "当日線" Then
            line.Delete
        End If
    Next
    
    ' 当日日付を取得
    Dim todayCell As Range
    Dim todayColumn As Integer
    Set todayCell = Rows(Range("_calendar").Row).Find(What:=DateValue(Date), LookIn:=xlFormulas)
    
    If todayCell Is Nothing Then
        MsgBox "当日日付がシート内に存在しません。" & vbCrLf _
        & "対象日付：" & Date
        Exit Sub
    End If
    
    todayColumn = todayCell.Column
    
    ' 線を出力させる座標を取得
    Dim startX As Double, startY As Double  ' 開始座標
    Dim endX As Double, endY As Double      ' 終了座標
    startX = Cells(Range("_calendar").Row, todayColumn).Left + 12
    startY = Cells(Range("_calendar").Row + 3, todayColumn).Top
    endX = Cells(Range("_calendar").Row, todayColumn).Left + 12
    endY = Cells(Range("_taskStart").End(xlDown).Row + 1, 1).Top

    '当日線を出力
    Dim ffb As FreeformBuilder
    Set ffb = ActiveSheet.Shapes.BuildFreeform(msoEditingCorner, startX, startY)
    ffb.AddNodes msoSegmentLine, msoEditingCorner, endX, endY
    
    Dim progress As Shape                   ' イナズマ線
    Set progress = ffb.ConvertToShape

    '当日線の書式編集
    With progress
        .Name = "当日線"
        .line.DashStyle = msoLineSolid       'スタイル
        .line.Weight = 1.5                   '太さ
        .line.ForeColor.RGB = RGB(255, 0, 0) '最新は赤
    End With

    Dim i As Integer
    Dim startDate As Date, endDate As Date, percentage As Double
    Dim startColumn As Integer, endColumn As Integer
    Dim startDateCell As Range, endDateCell As Range
    Dim vertexX As Double, vertexY As Double

    For i = Range("_taskStart").Row To Range("_taskStart").End(xlDown).Row

        ' 非表示行の場合スキップ
        If Rows(i).Hidden Then
            GoTo Continue
        End If

        ' ステータスが"対応中"、"保留"でない場合、スキップする
        If Cells(i, Range("_status").Column) <> "対応中" And Cells(i, Range("_status").Column) <> "保留" Then
            GoTo Continue
        End If
        
        ' 開始日、終了日、進捗率を取得
        startDate = Cells(i, Range("_startDate").Column).Value
        endDate = Cells(i, Range("_endDate").Column).Value
        percentage = Cells(i, Range("_progress").Column).Value
        
        ' 開始日の列番号を取得
        Set startDateCell = Rows(Range("_calendar").Row).Find(What:=DateValue(startDate), LookIn:=xlFormulas)
        If startDateCell Is Nothing Then
            GoTo Continue
        End If
        startColumn = startDateCell.Column
        
        ' 終了日の列番号を取得
        Set endDateCell = Rows(Range("_calendar").Row).Find(What:=DateValue(endDate), LookIn:=xlFormulas)
        If endDateCell Is Nothing Then
            GoTo Continue
        End If
        endColumn = endDateCell.Column
        
        ' 稲妻線の頂点定義
        startY = Cells(i, endColumn).Top
        endY = Cells(i + 1, endColumn).Top
        vertexX = Cells(i, startColumn).Left _
            + ((Cells(i, endColumn + 1).Left - Cells(i, startColumn).Left) _
                * percentage)
        vertexY = Cells(i, startColumn).Top + 4.375
                
        ' 頂点を描画
         With progress
             .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, startX, startY
             .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, vertexX, vertexY
             .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, endX, endY
         End With

Continue:
    Next i

    Application.ScreenUpdating = True

End Sub

Sub 当日線を描画する()

    ' -------------------------------------------------------------------------
    ' 当日線をカレンダー該当箇所に作成
    ' -------------------------------------------------------------------------
    Dim todayColumn As Integer
    Dim toDay As Date
    Dim line As Shape
    Dim startX As Double, startY As Double  ' 開始座標
    Dim endX As Double, endY As Double      ' 終了座標
    Dim foundCell As Range
    Dim todayLine As Integer
    
    '既存の当日線を削除
    Worksheets("WBS").Activate
    For Each line In ActiveSheet.Shapes
        'エクセルシート上で"当日線"を削除する
        If line.Name = "当日線" Then
            line.Delete
        End If
    Next
    
    ' 当日日付を取得
    toDay = Date
    Set foundCell = Rows(Range("_calendar").Row).Find(What:=DateValue(toDay), LookIn:=xlFormulas)
    
    If foundCell Is Nothing Then
        MsgBox "当日日付がシート内に存在しません。" & vbCrLf _
            & "対象日付：" & toDay
        Exit Sub
    End If
    
    todayColumn = foundCell.Column
    
    '線を出力させる座標を取得
    startX = Cells(Range("_calendar").Row, todayColumn).Left + 12
    startY = Cells(Range("_calendar").Row + 3, todayColumn).Top
    endX = Cells(Range("_calendar").Row, todayColumn).Left + 12
    endY = Cells(Range("_taskStart").End(xlDown).Row + 1, 1).Top
    
    '当日線を作成
    Set line = ActiveSheet.Shapes.AddLine(startX, startY, endX, endY)

    With line
        .Name = "当日線"
        .line.ForeColor.RGB = vbRed
        .line.Weight = 1.5
    End With
     
    foundCell.Activate

End Sub


Sub シートの条件付き書式を初期化する()

    Worksheets("WBS").Cells.FormatConditions.Delete
    
    Call カレンダーの条件付き書式を作成する
    Call テーブルの条件付き書式を作成する

End Sub

Public Sub カレンダーの条件付き書式を作成する()

    Dim calendarRange As Range
    Dim calendarRangeCondition As FormatCondition
    
    Worksheets("WBS").Activate
    Set calendarRange = Range( _
        Range("_calendar").Offset(3, 0), _
        Cells(Range("_taskStart").End(xlDown).Row, Range("_calendar").End(xlToRight).Column) _
    )
    
    ' -------------------------------------------------------------------------
    ' 格子書式を設定する
    ' -------------------------------------------------------------------------
    calendarRange.Borders(xlDiagonalDown).LineStyle = xlNone
    calendarRange.Borders(xlDiagonalUp).LineStyle = xlNone
    calendarRange.Borders(xlEdgeLeft).LineStyle = xlNone
    calendarRange.Borders(xlEdgeTop).LineStyle = xlNone
    With calendarRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThin
    End With
    calendarRange.Borders(xlEdgeRight).LineStyle = xlNone
    With calendarRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With calendarRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' -------------------------------------------------------------------------
    ' 週末セルの書式を設定する
    ' -------------------------------------------------------------------------
    Dim dateCell As String
    ' OR(WEEKDAY(K$5)=1, WEEKDAY(K$5)=7)
    dateCell = Range("_calendar").Address(True, False)
    Set calendarRangeCondition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=OR(WEEKDAY(" & dateCell & ")=1, WEEKDAY(" & dateCell & ")=7)" _
    )
    calendarRangeCondition.Interior.ColorIndex = 16
    
    ' -------------------------------------------------------------------------
    ' 祝日セルの書式を設定する
    ' -------------------------------------------------------------------------
    Set calendarRangeCondition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=COUNTIF(INDIRECT(""_祝日[日付]"")," & dateCell & ")=1" _
    )
    calendarRangeCondition.Interior.ColorIndex = 16

    ' -------------------------------------------------------------------------
    ' タスクラインの書式設定をする
    ' -------------------------------------------------------------------------
    ' =AND(K$5>=$G7,K$5<=$H7,$F7="終了")
    ' 開始日 < 日付 < 終了日の場合、かつステータスが対応完了の場合、グレーアウト
    Dim startDate, endDate, status As String
    startDate = Range("_startDate").Offset(1, 0).Address(False, True)
    endDate = Range("_endDate").Offset(1, 0).Address(False, True)
    status = Range("_status").Offset(1, 0).Address(False, True)
    
    Set calendarRangeCondition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND(" & dateCell & ">=" & startDate & "," & dateCell & "<=" & endDate & "," & status & "=""対応完了"")" _
    )
    calendarRangeCondition.Interior.ColorIndex = 15
    
    ' 開始日 < 日付 < 終了日の場合、かつ終了日 < 当日日付の場合、赤塗
    ' =AND(K$5>=$G7,K$5<=$H7,$H7<TODAY())
    Set calendarRangeCondition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND(" & dateCell & ">=" & startDate & "," & dateCell & "<=" & endDate & "," & endDate & "<TODAY())" _
    )
    calendarRangeCondition.Interior.ColorIndex = 3
    
    ' 開始日 < 日付 < 終了日の場合、緑塗
    ' =AND(K$5>=$G7,K$5<=$H7)
    Set calendarRangeCondition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND(" & dateCell & ">=" & startDate & "," & dateCell & "<=" & endDate & ")" _
    )
    calendarRangeCondition.Interior.ColorIndex = 35

End Sub

Public Sub テーブルの条件付き書式を作成する()

    Dim tableRange As Range

    ' -------------------------------------------------------------------------
    ' 書式を初期化する
    ' -------------------------------------------------------------------------
    ' 書式範囲取得
    Set tableRange = Range("WBS")
    
    ' -------------------------------------------------------------------------
    ' アイコン書式を設定する
    ' -------------------------------------------------------------------------
    Dim iconRange As Range
    Dim iconRangeCondition As IconSetCondition
    
    Set iconRange = Range( _
        Range("_progress").Offset(1, 0), _
        Cells(Range("_taskStart").End(xlDown).Row, Range("_progress").Column) _
    )
    Set iconRangeCondition = iconRange.FormatConditions.AddIconSetCondition()
    With iconRangeCondition
        .IconSet = ActiveWorkbook.IconSets(xl5Quarters)
        
        With .IconCriteria(2)
            .Type = xlConditionValueFormula
            .Value = "=0.25"
            .Operator = xlGreaterEqual
        End With
        With .IconCriteria(3)
            .Type = xlConditionValueFormula
            .Value = "=0.5"
            .Operator = xlGreaterEqual
        End With
        With .IconCriteria(4)
            .Type = xlConditionValueFormula
            .Value = "=0.75"
            .Operator = xlGreaterEqual
        End With
        With .IconCriteria(5)
            .Type = xlConditionValueFormula
            .Value = "=1"
            .Operator = xlGreaterEqual
        End With
    End With
    
    ' -------------------------------------------------------------------------
    ' 行書式を設定する
    ' -------------------------------------------------------------------------
    ' ステータスが対応完了の場合、行をグレーアウト
    Dim statusCell As String
    statusCell = Range("_status").Offset(1, 0).Address(False, True)
    
    Dim tableRangeCondition As FormatCondition
    Set tableRangeCondition = tableRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=" & statusCell & "=""対応完了""" _
    )
    tableRangeCondition.Interior.ColorIndex = 15
    
    ' 終了日 < 当日日付の場合、行を赤塗
    Dim endDateCell As String
    endDateCell = Range("_endDate").Offset(1, 0).Address(False, True)
    Set tableRangeCondition = tableRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND(" & endDateCell & "<TODAY()," & endDateCell & "<>"""")" _
    )
    tableRangeCondition.Interior.ColorIndex = 3
    
End Sub

Sub セルのメモを更新する()

    Application.ScreenUpdating = False
    Worksheets("WBS").Activate
    
    ' メモを初期化
On Error GoTo NoteError
    Dim noteCellRange As Range
    Set noteCellRange = Cells.SpecialCells(xlCellTypeComments)
    If Not noteCellRange Is Nothing Then
        noteCellRange.ClearComments
    End If
    
NoteError:
    Err.Clear
    
    Dim i As Integer
    Dim endDateCell As Range
    Dim endColumn As Integer
    Dim endDate As Date
    Dim endDateColumn As Integer, categoryColumn As Integer, subCategoryColumn As Integer, subSubCategoryColumn As Integer
    Dim category As String, subCategory As String, subSubCategory As String
    Dim noteMessage As String
    Dim comment As comment
    
    endDateColumn = Range("_endDate").Column
    categoryColumn = Range("_category").Column
    subCategoryColumn = Range("_subCategory").Column
    subSubCategoryColumn = Range("_subSubCategory").Column
    
    For i = Range("_taskStart").Row To Range("_taskStart").End(xlDown).Row
    
        ' 終了日、大分類、中分類、小分類を取得する
        endDate = Cells(i, endDateColumn).Value
        category = Cells(i, categoryColumn).Value
        subCategory = Cells(i, subCategoryColumn).Value
        subSubCategory = Cells(i, subSubCategoryColumn).Value
        
        ' メモに入力する文字列を作成する
        noteMessage = "[" & category & "]"
        If subCategory <> "" Then
            noteMessage = noteMessage & "：[" & subCategory & "]"
        End If
        If subSubCategory <> "" Then
            noteMessage = noteMessage & "：[" & subSubCategory & "]"
        End If
        
        ' メモを追加するセルを取得する
        Set endDateCell = Rows(Range("_calendar").Row).Find(What:=DateValue(endDate), LookIn:=xlFormulas)
        If endDateCell Is Nothing Then
            GoTo Continue
        End If
        endColumn = endDateCell.Column
        
        If noteMessage <> "[]" Then
            Set comment = Cells(i, endColumn).AddComment(noteMessage)
            With comment.Shape
                .TextFrame.AutoSize = True
                .TextFrame.Characters.Font.Name = "Arial"
                .TextFrame.Characters.Font.ColorIndex = 1
                .TextFrame.Characters.Font.Size = 11
                .Fill.ForeColor.SchemeColor = 80
                .Placement = xlMoveAndSize
            End With
        End If
    
Continue:
    Next i
    
    Application.ScreenUpdating = True

End Sub

Sub セルコメントの表示非表示を切り替える()

    If Application.DisplayCommentIndicator = xlCommentIndicatorOnly Then
        Application.DisplayCommentIndicator = xlCommentAndIndicator
    Else
        Application.DisplayCommentIndicator = xlCommentIndicatorOnly
    End If

End Sub

Sub WSBを初期化する()

    Application.ScreenUpdating = False

    Worksheets("WBS").Activate
    
    ' 開始日・終了日を取得
    Dim startDate As String
    Dim endDate As String
    
InputData:
    startDate = InputBox("開始日を入力してください。(yyyy/MM/dd)", "カレンダー設定")
    If StrPtr(startDate) = 0 Then
        Exit Sub
    ElseIf Not IsDate(startDate) Then
        MsgBox "yyyy/MM/dd形式で入力してください。"
        GoTo InputData
    End If
    
    endDate = InputBox("終了日を入力してください。(yyyy/MM/dd)", "カレンダー設定")
    If StrPtr(endDate) = 0 Then
        Exit Sub
    ElseIf Not IsDate(endDate) Then
        MsgBox "yyyy/MM/dd形式で入力してください。"
        GoTo InputData
    End If
    
    Dim calendar As Range
    Set calendar = Range("_calendar")
    Dim dateRange As Range
    
    ' カレンダー範囲を初期化
    Set dateRange = Range( _
        calendar.Offset(-1, 0), _
        calendar.End(xlToRight).Offset(1, 0) _
    )
    dateRange.Clear
        
    ' カレンダーを生成
    Dim diff As Integer
    diff = DateDiff("d", DateValue(startDate), DateValue(endDate))
    calendar.Value = startDate
    
    Dim dayRange As Range
    Dim monthRange As Range
    Dim weekRange As Range
    Set dayRange = Range(calendar, calendar.Offset(0, diff))
    Set monthRange = Range(calendar.Offset(-1, 0), calendar.Offset(-1, diff))
    Set weekRange = Range(calendar.Offset(1, 0), calendar.Offset(1, diff))
    
    calendar.AutoFill _
        dayRange, _
        xlFillDays
    
    dayRange.Copy monthRange
    dayRange.Copy weekRange
    
    ' 書式設定
    monthRange.NumberFormat = "m"       ' 月
    dayRange.NumberFormat = "d"         ' 日
    weekRange.NumberFormat = "aaa"      ' 曜日

    Set dateRange = Range( _
        calendar.Offset(-1, 0), _
        calendar.End(xlToRight).Offset(1, 0) _
    )
    dateRange.Interior.Color = Range("_startDate").Offset(-1, 0).Interior.Color
    dateRange.Font.ColorIndex = 2
    dateRange.Font.Bold = True
    
    Call 当日線を描画する
    Call 実績線を描画する
    Call セルのメモを更新する
    Call シートの条件付き書式を初期化する

    Application.ScreenUpdating = True

End Sub
