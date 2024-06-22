Attribute VB_Name = "GanttChart"

Public Sub ���ѐ���`�悷��()

    Application.ScreenUpdating = False

    Dim startDate, endDate As Date                      ' �J�n��, �I����
    Dim percentage As Double                            ' �i����
    Dim square As Shape                                 ' ���ѐ������`
    Dim leftPoint, topPoint, height, width As Double    ' ���ѐ��T�C�Y
    Dim statusColumn, startDateColumn, endDateColumn, percentageColumn As Integer
    Dim status As String
    Dim i As Integer
    Dim startDateCell As Range, endDateCell As Range
    
    ' ���ѐ����폜����
    Worksheets("WBS").Activate
    For Each square In ActiveSheet.Shapes
        If square.Name = "���ѐ�" Then
            square.Delete
        End If
    Next
    
    statusColumn = Range("_status").Column              ' �� ��
    startDateColumn = Range("_startDate").Column        ' �J�n�� ��
    endDateColumn = Range("_endDate").Column            ' �I���� ��
    percentageColumn = Range("_progress").Column        ' �i���� ��
    
    For i = Range("_taskStart").Row To Range("_taskStart").End(xlDown).Row

        ' ��\���s�̏ꍇ�X�L�b�v
        If Rows(i).Hidden Then
            GoTo Continue
        End If

        ' �X�e�[�^�X���Ή������A�Ή����A�ۗ��łȂ��ꍇ
        status = Cells(i, statusColumn).Value
        If status <> "�Ή�����" And status <> "�Ή���" And status <> "�ۗ�" Then
            GoTo Continue
        End If
        
        ' �J�n���A�I�����A�i�������擾
        startDate = Cells(i, startDateColumn).Value
        endDate = Cells(i, endDateColumn).Value
        percentage = Cells(i, percentageColumn).Value
        
        ' �J�����_�[�s����J�n���Z���A�I�����Z�����擾
        Set startDateCell = Rows(Range("_calendar").Row).Find(What:=DateValue(startDate), LookIn:=xlFormulas)
        Set endDateCell = Rows(Range("_calendar").Row).Find(What:=DateValue(endDate), LookIn:=xlFormulas)
        If startDateCell Is Nothing Or endDateCell Is Nothing Then
            GoTo Continue
        End If
        
        Dim startColumn As Integer, endColumn As Integer
        startColumn = startDateCell.Column
        endColumn = endDateCell.Column
        
        ' ���ѐ��}�`�̃T�C�Y�E�ʒu�ݒ�
        leftPoint = Cells(i, startColumn).Left                                                  ' ���ʒu
        topPoint = Cells(i, startColumn).Top                                                    ' ��ʒu
        height = Cells(i, startColumn).height - 10                                              ' ����
        width = (Cells(i, endColumn + 1).Left - Cells(i, startColumn).Left) * percentage        ' ��
        
        ' ���ѐ����쐬
        Set square = ActiveSheet.Shapes.AddShape(msoShapeRectangle, leftPoint, topPoint, width, height)
        square.Fill.ForeColor.SchemeColor = 32              ' �F�F��
        square.Name = "���ѐ�"                              ' ����
        square.Fill.Transparency = 0.4                      ' �����x�F40%
        
Continue:
    Next i
    
    Application.ScreenUpdating = True

End Sub


Sub ��Ȑ���`�悷��()

    Application.ScreenUpdating = False

    Dim line As Shape
    
    Worksheets("WBS").Activate

    ' �����̈�Ȑ����폜
    For Each line In ActiveSheet.Shapes
        If line.Name = "������" Then
            line.Delete
        End If
    Next
    
    ' �������t���擾
    Dim todayCell As Range
    Dim todayColumn As Integer
    Set todayCell = Rows(Range("_calendar").Row).Find(What:=DateValue(Date), LookIn:=xlFormulas)
    
    If todayCell Is Nothing Then
        MsgBox "�������t���V�[�g���ɑ��݂��܂���B" & vbCrLf _
        & "�Ώۓ��t�F" & Date
        Exit Sub
    End If
    
    todayColumn = todayCell.Column
    
    ' �����o�͂�������W���擾
    Dim startX As Double, startY As Double  ' �J�n���W
    Dim endX As Double, endY As Double      ' �I�����W
    startX = Cells(Range("_calendar").Row, todayColumn).Left + 12
    startY = Cells(Range("_calendar").Row + 3, todayColumn).Top
    endX = Cells(Range("_calendar").Row, todayColumn).Left + 12
    endY = Cells(Range("_taskStart").End(xlDown).Row + 1, 1).Top

    '���������o��
    Dim ffb As FreeformBuilder
    Set ffb = ActiveSheet.Shapes.BuildFreeform(msoEditingCorner, startX, startY)
    ffb.AddNodes msoSegmentLine, msoEditingCorner, endX, endY
    
    Dim progress As Shape                   ' �C�i�Y�}��
    Set progress = ffb.ConvertToShape

    '�������̏����ҏW
    With progress
        .Name = "������"
        .line.DashStyle = msoLineSolid       '�X�^�C��
        .line.Weight = 1.5                   '����
        .line.ForeColor.RGB = RGB(255, 0, 0) '�ŐV�͐�
    End With

    Dim i As Integer
    Dim startDate As Date, endDate As Date, percentage As Double
    Dim startColumn As Integer, endColumn As Integer
    Dim startDateCell As Range, endDateCell As Range
    Dim vertexX As Double, vertexY As Double

    For i = Range("_taskStart").Row To Range("_taskStart").End(xlDown).Row

        ' ��\���s�̏ꍇ�X�L�b�v
        If Rows(i).Hidden Then
            GoTo Continue
        End If

        ' �X�e�[�^�X��"�Ή���"�A"�ۗ�"�łȂ��ꍇ�A�X�L�b�v����
        If Cells(i, Range("_status").Column) <> "�Ή���" And Cells(i, Range("_status").Column) <> "�ۗ�" Then
            GoTo Continue
        End If
        
        ' �J�n���A�I�����A�i�������擾
        startDate = Cells(i, Range("_startDate").Column).Value
        endDate = Cells(i, Range("_endDate").Column).Value
        percentage = Cells(i, Range("_progress").Column).Value
        
        ' �J�n���̗�ԍ����擾
        Set startDateCell = Rows(Range("_calendar").Row).Find(What:=DateValue(startDate), LookIn:=xlFormulas)
        If startDateCell Is Nothing Then
            GoTo Continue
        End If
        startColumn = startDateCell.Column
        
        ' �I�����̗�ԍ����擾
        Set endDateCell = Rows(Range("_calendar").Row).Find(What:=DateValue(endDate), LookIn:=xlFormulas)
        If endDateCell Is Nothing Then
            GoTo Continue
        End If
        endColumn = endDateCell.Column
        
        ' ��Ȑ��̒��_��`
        startY = Cells(i, endColumn).Top
        endY = Cells(i + 1, endColumn).Top
        vertexX = Cells(i, startColumn).Left _
            + ((Cells(i, endColumn + 1).Left - Cells(i, startColumn).Left) _
                * percentage)
        vertexY = Cells(i, startColumn).Top + 4.375
                
        ' ���_��`��
         With progress
             .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, startX, startY
             .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, vertexX, vertexY
             .Nodes.Insert .Nodes.Count - 1, msoSegmentLine, msoEditingAuto, endX, endY
         End With

Continue:
    Next i

    Application.ScreenUpdating = True

End Sub

Sub ��������`�悷��()

    ' -------------------------------------------------------------------------
    ' ���������J�����_�[�Y���ӏ��ɍ쐬
    ' -------------------------------------------------------------------------
    Dim todayColumn As Integer
    Dim toDay As Date
    Dim line As Shape
    Dim startX As Double, startY As Double  ' �J�n���W
    Dim endX As Double, endY As Double      ' �I�����W
    Dim foundCell As Range
    Dim todayLine As Integer
    
    '�����̓��������폜
    Worksheets("WBS").Activate
    For Each line In ActiveSheet.Shapes
        '�G�N�Z���V�[�g���"������"���폜����
        If line.Name = "������" Then
            line.Delete
        End If
    Next
    
    ' �������t���擾
    toDay = Date
    Set foundCell = Rows(Range("_calendar").Row).Find(What:=DateValue(toDay), LookIn:=xlFormulas)
    
    If foundCell Is Nothing Then
        MsgBox "�������t���V�[�g���ɑ��݂��܂���B" & vbCrLf _
            & "�Ώۓ��t�F" & toDay
        Exit Sub
    End If
    
    todayColumn = foundCell.Column
    
    '�����o�͂�������W���擾
    startX = Cells(Range("_calendar").Row, todayColumn).Left + 12
    startY = Cells(Range("_calendar").Row + 3, todayColumn).Top
    endX = Cells(Range("_calendar").Row, todayColumn).Left + 12
    endY = Cells(Range("_taskStart").End(xlDown).Row + 1, 1).Top
    
    '���������쐬
    Set line = ActiveSheet.Shapes.AddLine(startX, startY, endX, endY)

    With line
        .Name = "������"
        .line.ForeColor.RGB = vbRed
        .line.Weight = 1.5
    End With
     
    foundCell.Activate

End Sub


Sub �V�[�g�̏����t������������������()

    Worksheets("WBS").Cells.FormatConditions.Delete
    
    Call �J�����_�[�̏����t���������쐬����
    Call �e�[�u���̏����t���������쐬����

End Sub

Public Sub �J�����_�[�̏����t���������쐬����()

    Dim calendarRange As Range
    Dim calendarRangeCondition As FormatCondition
    
    Worksheets("WBS").Activate
    Set calendarRange = Range( _
        Range("_calendar").Offset(3, 0), _
        Cells(Range("_taskStart").End(xlDown).Row, Range("_calendar").End(xlToRight).Column) _
    )
    
    ' -------------------------------------------------------------------------
    ' �i�q������ݒ肷��
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
    ' �T���Z���̏�����ݒ肷��
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
    ' �j���Z���̏�����ݒ肷��
    ' -------------------------------------------------------------------------
    Set calendarRangeCondition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=COUNTIF(INDIRECT(""_�j��[���t]"")," & dateCell & ")=1" _
    )
    calendarRangeCondition.Interior.ColorIndex = 16

    ' -------------------------------------------------------------------------
    ' �^�X�N���C���̏����ݒ������
    ' -------------------------------------------------------------------------
    ' =AND(K$5>=$G7,K$5<=$H7,$F7="�I��")
    ' �J�n�� < ���t < �I�����̏ꍇ�A���X�e�[�^�X���Ή������̏ꍇ�A�O���[�A�E�g
    Dim startDate, endDate, status As String
    startDate = Range("_startDate").Offset(1, 0).Address(False, True)
    endDate = Range("_endDate").Offset(1, 0).Address(False, True)
    status = Range("_status").Offset(1, 0).Address(False, True)
    
    Set calendarRangeCondition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND(" & dateCell & ">=" & startDate & "," & dateCell & "<=" & endDate & "," & status & "=""�Ή�����"")" _
    )
    calendarRangeCondition.Interior.ColorIndex = 15
    
    ' �J�n�� < ���t < �I�����̏ꍇ�A���I���� < �������t�̏ꍇ�A�ԓh
    ' =AND(K$5>=$G7,K$5<=$H7,$H7<TODAY())
    Set calendarRangeCondition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND(" & dateCell & ">=" & startDate & "," & dateCell & "<=" & endDate & "," & endDate & "<TODAY())" _
    )
    calendarRangeCondition.Interior.ColorIndex = 3
    
    ' �J�n�� < ���t < �I�����̏ꍇ�A�Γh
    ' =AND(K$5>=$G7,K$5<=$H7)
    Set calendarRangeCondition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND(" & dateCell & ">=" & startDate & "," & dateCell & "<=" & endDate & ")" _
    )
    calendarRangeCondition.Interior.ColorIndex = 35

End Sub

Public Sub �e�[�u���̏����t���������쐬����()

    Dim tableRange As Range

    ' -------------------------------------------------------------------------
    ' ����������������
    ' -------------------------------------------------------------------------
    ' �����͈͎擾
    Set tableRange = Range("WBS")
    
    ' -------------------------------------------------------------------------
    ' �A�C�R��������ݒ肷��
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
    ' �s������ݒ肷��
    ' -------------------------------------------------------------------------
    ' �X�e�[�^�X���Ή������̏ꍇ�A�s���O���[�A�E�g
    Dim statusCell As String
    statusCell = Range("_status").Offset(1, 0).Address(False, True)
    
    Dim tableRangeCondition As FormatCondition
    Set tableRangeCondition = tableRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=" & statusCell & "=""�Ή�����""" _
    )
    tableRangeCondition.Interior.ColorIndex = 15
    
    ' �I���� < �������t�̏ꍇ�A�s��ԓh
    Dim endDateCell As String
    endDateCell = Range("_endDate").Offset(1, 0).Address(False, True)
    Set tableRangeCondition = tableRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND(" & endDateCell & "<TODAY()," & endDateCell & "<>"""")" _
    )
    tableRangeCondition.Interior.ColorIndex = 3
    
End Sub

Sub �Z���̃������X�V����()

    Application.ScreenUpdating = False
    Worksheets("WBS").Activate
    
    ' ������������
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
    
        ' �I�����A�啪�ށA�����ށA�����ނ��擾����
        endDate = Cells(i, endDateColumn).Value
        category = Cells(i, categoryColumn).Value
        subCategory = Cells(i, subCategoryColumn).Value
        subSubCategory = Cells(i, subSubCategoryColumn).Value
        
        ' �����ɓ��͂��镶������쐬����
        noteMessage = "[" & category & "]"
        If subCategory <> "" Then
            noteMessage = noteMessage & "�F[" & subCategory & "]"
        End If
        If subSubCategory <> "" Then
            noteMessage = noteMessage & "�F[" & subSubCategory & "]"
        End If
        
        ' ������ǉ�����Z�����擾����
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

Sub �Z���R�����g�̕\����\����؂�ւ���()

    If Application.DisplayCommentIndicator = xlCommentIndicatorOnly Then
        Application.DisplayCommentIndicator = xlCommentAndIndicator
    Else
        Application.DisplayCommentIndicator = xlCommentIndicatorOnly
    End If

End Sub

Sub WSB������������()

    Application.ScreenUpdating = False

    Worksheets("WBS").Activate
    
    ' �J�n���E�I�������擾
    Dim startDate As String
    Dim endDate As String
    
InputData:
    startDate = InputBox("�J�n������͂��Ă��������B(yyyy/MM/dd)", "�J�����_�[�ݒ�")
    If StrPtr(startDate) = 0 Then
        Exit Sub
    ElseIf Not IsDate(startDate) Then
        MsgBox "yyyy/MM/dd�`���œ��͂��Ă��������B"
        GoTo InputData
    End If
    
    endDate = InputBox("�I��������͂��Ă��������B(yyyy/MM/dd)", "�J�����_�[�ݒ�")
    If StrPtr(endDate) = 0 Then
        Exit Sub
    ElseIf Not IsDate(endDate) Then
        MsgBox "yyyy/MM/dd�`���œ��͂��Ă��������B"
        GoTo InputData
    End If
    
    Dim calendar As Range
    Set calendar = Range("_calendar")
    Dim dateRange As Range
    
    ' �J�����_�[�͈͂�������
    Set dateRange = Range( _
        calendar.Offset(-1, 0), _
        calendar.End(xlToRight).Offset(1, 0) _
    )
    dateRange.Clear
        
    ' �J�����_�[�𐶐�
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
    
    ' �����ݒ�
    monthRange.NumberFormat = "m"       ' ��
    dayRange.NumberFormat = "d"         ' ��
    weekRange.NumberFormat = "aaa"      ' �j��

    Set dateRange = Range( _
        calendar.Offset(-1, 0), _
        calendar.End(xlToRight).Offset(1, 0) _
    )
    dateRange.Interior.Color = Range("_startDate").Offset(-1, 0).Interior.Color
    dateRange.Font.ColorIndex = 2
    dateRange.Font.Bold = True
    
    Call ��������`�悷��
    Call ���ѐ���`�悷��
    Call �Z���̃������X�V����
    Call �V�[�g�̏����t������������������

    Application.ScreenUpdating = True

End Sub
