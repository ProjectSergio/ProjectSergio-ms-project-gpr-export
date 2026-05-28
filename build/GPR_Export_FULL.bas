Option Explicit

Private isLoading As Boolean
Private ExportColWidths(0 To 32) As Long
Private ExportColWidthsInitialized As Boolean

Private Type TMonthSummary
    StartDate As Date
    EndDate As Date
    Label As String
    YearNum As Long
    DaysInPeriod As Long

    BaseTotal As Double
    BaseCMR As Double
    BaseEquip As Double

    PlanTotal As Double
    PlanCMR As Double
    PlanEquip As Double

    FactTotal As Double
    FactCMR As Double
    FactEquip As Double

    TzrSmeta As Double
    TzrSmetaSub As Double
    TzrBudgetSS As Double
    PeopleSub As Double
    PeopleSSPlanTzr As Double
    TzrFactSS As Double

    MimSmeta As Double
    MimBudget As Double
    MimFact As Double

    PlanCum As Double
    FactCum As Double
    BaseCum As Double

    PercentCompleteMoney As Double
    PercentExec As Double
End Type

Private Type TPeriodResourceValues
    HumanPlan As Double
    HumanFact As Double
    MimPlan As Double
    MimFact As Double
    VolumePlan As Double
    VolumeFact As Double
    VolumeFactCum As Double
    VolumeFactPast As Double
End Type

' ============================================================
' ДАТЫ / ФОРМА
' ============================================================

Private Function TryParseDate(ByVal s As String, ByRef outDate As Date) As Boolean
    On Error GoTo Fallback

    outDate = CDate(s)
    TryParseDate = True
    Exit Function

Fallback:
    On Error GoTo 0

    Dim p() As String
    p = Split(s, ".")

    If UBound(p) = 2 Then
        If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
            outDate = DateSerial(CLng(p(2)), CLng(p(1)), CLng(p(0)))
            TryParseDate = True
            Exit Function
        End If
    End If

    TryParseDate = False
End Function

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Initialize()
    isLoading = True

    Dim dStart As Date
    Dim dEnd As Date

    dStart = DateSerial(Year(Date), Month(Date), 1)
    dEnd = DateSerial(Year(Date), Month(Date) + 1, 0)

    InitYearMonth dStart

    TextBox1.Value = Format$(dStart, "dd.mm.yyyy")
    TextBox2.Value = Format$(dEnd, "dd.mm.yyyy")

    isLoading = False
End Sub

Private Sub InitYearMonth(ByVal d As Date)
    Dim i As Long

    ComboBox1.Clear
    For i = Year(Date) - 5 To Year(Date) + 5
        ComboBox1.AddItem CStr(i)
    Next i
    ComboBox1.Value = CStr(Year(d))

    ComboBox2.Clear
    For i = 1 To 12
        ComboBox2.AddItem Format(DateSerial(2000, i, 1), "mmmm")
    Next i
    ComboBox2.ListIndex = Month(d) - 1
End Sub

Private Sub SyncTextBoxesFromYM()
    If ComboBox1.ListIndex = -1 Or ComboBox2.ListIndex = -1 Then Exit Sub

    Dim y As Long
    Dim m As Long
    Dim dStart As Date
    Dim dEnd As Date

    y = CLng(ComboBox1.Value)
    m = ComboBox2.ListIndex + 1

    dStart = DateSerial(y, m, 1)
    dEnd = DateSerial(y, m + 1, 0)

    TextBox1.Value = Format$(dStart, "dd.mm.yyyy")
    TextBox2.Value = Format$(dEnd, "dd.mm.yyyy")
End Sub

Private Sub SyncYMFromTextBoxes()
    Dim d As Date

    If TryParseDate(TextBox1.Text, d) Then
        isLoading = True
        ComboBox1.Value = CStr(Year(d))
        ComboBox2.ListIndex = Month(d) - 1
        isLoading = False
    End If
End Sub

Private Sub ComboBox1_Change()
    If Not isLoading Then SyncTextBoxesFromYM
End Sub

Private Sub ComboBox2_Change()
    If Not isLoading Then SyncTextBoxesFromYM
End Sub

Private Sub TextBox1_AfterUpdate()
    If Not isLoading Then SyncYMFromTextBoxes
End Sub

Private Sub TextBox2_AfterUpdate()

End Sub

' ============================================================
' КНОПКИ ФОРМЫ
' ============================================================

Private Sub CommandButton1_Click()
    Dim originalCalculation As PjCalculation
    originalCalculation = Application.Calculation
    Application.Calculation = pjManual

    On Error GoTo EH

    Me.Hide

    Dim SD As Date
    Dim ED As Date

    SD = CDate(TextBox1.Value)
    ED = CDate(TextBox2.Value)

    If Int(ED) < Int(SD) Then
        MsgBox "Дата окончания меньше даты начала.", vbExclamation
        GoTo Cleanup
    End If

    RunCalcAndFilter SD, ED, True, True

Cleanup:
    Application.Calculation = originalCalculation
    Application.StatusBar = False
    Exit Sub

EH:
    MsgBox "Ошибка: " & Err.Number & " — " & Err.Description, vbExclamation
    Resume Cleanup
End Sub

Private Sub CommandButton3_Click()
    On Error GoTo EH

    isLoading = True

    Dim SD As Date
    Dim ED As Date

    SD = Int(ActiveProject.ProjectStart)
    ED = Int(ActiveProject.ProjectFinish)

    TextBox1.Value = Format$(SD, "dd.mm.yyyy")
    TextBox2.Value = Format$(ED, "dd.mm.yyyy")

    SyncYMFromTextBoxes

    isLoading = False
    Exit Sub

EH:
    isLoading = False
    MsgBox "Не удалось установить период всего проекта: " & Err.Description, vbExclamation
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

' ============================================================
' ТАБЛИЦА ДЛЯ ЭКСПОРТА
' ============================================================

Private Sub EnsureExportTable()
    Const EXPORT_TABLE_NAME As String = "ГПР экспорт (авто)"

    Dim fieldNames(0 To 32) As String
    Dim titles(0 To 32) As String
    Dim i As Long
    Dim tbl As Table

    On Error Resume Next
    Set tbl = ActiveProject.TaskTables(EXPORT_TABLE_NAME)
    On Error GoTo 0

    If Not tbl Is Nothing Then Exit Sub

    fieldNames(0) = "ID"
    fieldNames(1) = "Text2"
    fieldNames(2) = "Text7"
    fieldNames(3) = "Outline Level"
    fieldNames(4) = "Text4"
    fieldNames(5) = "Text10"
    fieldNames(6) = "Text20"
    fieldNames(7) = "Text21"
    fieldNames(8) = "Text22"
    fieldNames(9) = "Name"
    fieldNames(10) = "Duration"
    fieldNames(11) = "Start"
    fieldNames(12) = "Finish"
    fieldNames(13) = "Text1"
    fieldNames(14) = "Number1"
    fieldNames(15) = "Number2"
    fieldNames(16) = "Number3"
    fieldNames(17) = "Number11"
    fieldNames(18) = "Cost5"
    fieldNames(19) = "Cost1"
    fieldNames(20) = "Cost4"
    fieldNames(21) = "Cost8"
    fieldNames(22) = "Cost7"
    fieldNames(23) = "Resource Names"
    fieldNames(24) = "Number17"
    fieldNames(25) = "Number6"
    fieldNames(26) = "Number8"
    fieldNames(27) = "Number18"
    fieldNames(28) = "Number10"
    fieldNames(29) = "Number12"
    fieldNames(30) = "Cost2"
    fieldNames(31) = "Cost3"
    fieldNames(32) = "Cost6"

    titles(0) = "Ид."
    titles(1) = "№ п/п"
    titles(2) = "% Освоения ДС"
    titles(3) = "Уровень структуры"
    titles(4) = "Тип задачи"
    titles(5) = "Выполнение"
    titles(6) = "Объект"
    titles(7) = "Подобъект"
    titles(8) = "Раздел"
    titles(9) = "Название задачи"
    titles(10) = "Длительность"
    titles(11) = "Начало"
    titles(12) = "Окончание"
    titles(13) = "Ед. изм."
    titles(14) = "Объем проект"
    titles(15) = "Объем план"
    titles(16) = "Объем факт"
    titles(17) = "Объем факт (нак.)"
    titles(18) = "Стоимость См (руб. с НДС)"
    titles(19) = "Стоимость См план (с превыш.)"
    titles(20) = "Стоимость См факт (руб. с НДС)"
    titles(21) = "Стоимость См план (руб. с НДС)"
    titles(22) = "Стоимость См факт (без превыш.)"
    titles(23) = "Название ресурсов"
    titles(24) = "ТЗр Смета план (ч/ч)"
    titles(25) = "ТЗр Бюджет план (ч/ч)"
    titles(26) = "ТЗр факт (ч/ч)"
    titles(27) = "МиМ Смета план (м/ч)"
    titles(28) = "МиМ Бюджет план (м/ч)"
    titles(29) = "МиМ факт (м/ч)"
    titles(30) = "Затраты персонал"
    titles(31) = "Затраты МиМ"
    titles(32) = "Затраты персонал+МиМ"

    ExportColWidths(0) = 5
    ExportColWidths(1) = 8
    ExportColWidths(2) = 10
    ExportColWidths(3) = 10
    ExportColWidths(4) = 10
    ExportColWidths(5) = 12
    ExportColWidths(6) = 12
    ExportColWidths(7) = 12
    ExportColWidths(8) = 12
    ExportColWidths(9) = 30
    ExportColWidths(10) = 8
    ExportColWidths(11) = 12
    ExportColWidths(12) = 12
    ExportColWidths(13) = 8
    ExportColWidths(14) = 10
    ExportColWidths(15) = 10
    ExportColWidths(16) = 10
    ExportColWidths(17) = 12
    ExportColWidths(18) = 14
    ExportColWidths(19) = 14
    ExportColWidths(20) = 14
    ExportColWidths(21) = 14
    ExportColWidths(22) = 14
    ExportColWidths(23) = 20
    ExportColWidths(24) = 14
    ExportColWidths(25) = 14
    ExportColWidths(26) = 14
    ExportColWidths(27) = 14
    ExportColWidths(28) = 14
    ExportColWidths(29) = 14
    ExportColWidths(30) = 14
    ExportColWidths(31) = 14
    ExportColWidths(32) = 16

    ExportColWidthsInitialized = True

    On Error Resume Next

    Application.TableEditEx Name:=EXPORT_TABLE_NAME, TaskTable:=True, _
        Create:=True, OverwriteExisting:=True, _
        FieldName:=fieldNames(0), Title:=titles(0), Width:=ExportColWidths(0), _
        Align:=1, ShowInMenu:=False, LockFirstColumn:=True, _
        DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
        HeaderAutoRowHeightAdjustment:=False, WrapText:=False, _
        ShowAddNewColumn:=False

    For i = 1 To 32
        Application.TableEditEx Name:=EXPORT_TABLE_NAME, TaskTable:=True, _
            NewFieldName:=fieldNames(i), Title:=titles(i), Width:=ExportColWidths(i), _
            Align:=0, LockFirstColumn:=False, _
            DateFormat:=255, RowHeight:=1, AlignTitle:=1, _
            HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    Next i

    On Error GoTo 0
End Sub

Private Sub CaptureExportColumnWidthsFromProject()
    Const EXPORT_TABLE_NAME As String = "ГПР экспорт (авто)"

    On Error GoTo ExitHere

    Dim tbl As Table
    Dim i As Long
    Dim fCount As Long

    Set tbl = ActiveProject.TaskTables(EXPORT_TABLE_NAME)

    fCount = tbl.TableFields.Count
    If fCount > 33 Then fCount = 33

    For i = 1 To fCount
        ExportColWidths(i - 1) = tbl.TableFields(i).Width
    Next i

    ExportColWidthsInitialized = True

ExitHere:
    On Error GoTo 0
End Sub

' ============================================================
' ЭКСПОРТ В EXCEL — УСКОРЕННАЯ И СТАБИЛЬНАЯ ВЕРСИЯ
' ============================================================

Private Sub CommandButton4_Click()
    On Error GoTo EH

    Const EXPORT_TABLE_NAME As String = "ГПР экспорт (авто)"

    Dim originalProjectCalculation As PjCalculation
    Dim projectCalcSaved As Boolean

    Dim oldProjectScreenUpdating As Boolean
    Dim projectScreenSaved As Boolean

    Dim prevView As String
    Dim prevTable As String

    Dim SD As Date
    Dim ED As Date
    Dim SD0 As Date
    Dim ED0 As Date

    Dim xlApp As Object
    Dim wb As Object
    Dim wsSummary As Object

    Dim oldXlCalculation As Variant
    Dim oldXlCalculationSaved As Boolean

    Dim oldXlScreenUpdating As Boolean
    Dim oldXlScreenUpdatingSaved As Boolean

    Dim oldXlEnableEvents As Boolean
    Dim oldXlEnableEventsSaved As Boolean

    Dim oldXlDisplayAlerts As Boolean
    Dim oldXlDisplayAlertsSaved As Boolean

    Dim totalMonths As Long
    Dim madeMonthSheets As Long
    Dim monthData() As TMonthSummary

    On Error Resume Next

    Err.Clear
    originalProjectCalculation = Application.Calculation
    projectCalcSaved = (Err.Number = 0)

    If projectCalcSaved Then
        Err.Clear
        Application.Calculation = pjManual
    End If

    Err.Clear
    oldProjectScreenUpdating = Application.ScreenUpdating
    projectScreenSaved = (Err.Number = 0)

    If projectScreenSaved Then
        Err.Clear
        Application.ScreenUpdating = False
    End If

    Err.Clear
    If Not ActiveWindow Is Nothing Then prevView = ActiveWindow.View.Name

    Err.Clear
    prevTable = Application.CurrentTable

    On Error GoTo EH

    On Error Resume Next
    Application.ViewApply "Лист задач"

    If Err.Number <> 0 Then
        Err.Clear
        Application.ViewApply "Task Sheet"
    End If

    On Error GoTo EH

    EnsureExportTable

    On Error Resume Next
    Application.TableApply EXPORT_TABLE_NAME
    On Error GoTo EH

    CaptureExportColumnWidthsFromProject

    If Not TryParseDate(TextBox1.Value, SD) Then
        MsgBox "Некорректная дата начала.", vbExclamation
        GoTo Cleanup
    End If

    If Not TryParseDate(TextBox2.Value, ED) Then
        MsgBox "Некорректная дата окончания.", vbExclamation
        GoTo Cleanup
    End If

    SD0 = Int(SD)
    ED0 = Int(ED)

    If ED0 < SD0 Then
        MsgBox "Дата окончания меньше даты начала.", vbExclamation
        GoTo Cleanup
    End If

    totalMonths = CountMonthsInclusive(SD0, ED0)

    If totalMonths <= 0 Then
        MsgBox "Не удалось определить месяцы выгрузки.", vbExclamation
        GoTo Cleanup
    End If

    ReDim monthData(1 To totalMonths)

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True

    SaveAndOptimizeExcelState xlApp, _
                              oldXlCalculation, oldXlCalculationSaved, _
                              oldXlScreenUpdating, oldXlScreenUpdatingSaved, _
                              oldXlEnableEvents, oldXlEnableEventsSaved, _
                              oldXlDisplayAlerts, oldXlDisplayAlertsSaved

    Set wb = xlApp.Workbooks.Add

    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop

    Set wsSummary = wb.Worksheets(1)
    wsSummary.Name = GetUniqueSheetName(wb, "Выгрузка ГПР")

    madeMonthSheets = ExportPeriodByMonthsToWorkbook(wb, xlApp, SD0, ED0, monthData, totalMonths)

    BuildSummarySheet wsSummary, xlApp, monthData, madeMonthSheets, SD0, ED0

    wsSummary.Activate

    On Error Resume Next
    xlApp.Calculate
    On Error GoTo EH

    MsgBox "Экспорт завершён." & vbCrLf & _
           "Период: " & Format(SD0, "dd.mm.yyyy") & " — " & Format(ED0, "dd.mm.yyyy") & vbCrLf & _
           "Создан лист всего периода и месячных листов: " & madeMonthSheets & vbCrLf & _
           "Сводный лист и диаграммы сформированы.", vbInformation

Cleanup:
    On Error Resume Next

    RestoreExcelStateSafe xlApp, _
                          oldXlCalculation, oldXlCalculationSaved, _
                          oldXlScreenUpdating, oldXlScreenUpdatingSaved, _
                          oldXlEnableEvents, oldXlEnableEventsSaved, _
                          oldXlDisplayAlerts, oldXlDisplayAlertsSaved

    ResetViewFilter

    If Len(prevTable) > 0 Then Application.TableApply prevTable
    If Len(prevView) > 0 Then Application.ViewApply prevView

    If projectScreenSaved Then Application.ScreenUpdating = oldProjectScreenUpdating
    If projectCalcSaved Then Application.Calculation = originalProjectCalculation

    Application.StatusBar = False

    On Error GoTo 0
    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next

    RestoreExcelStateSafe xlApp, _
                          oldXlCalculation, oldXlCalculationSaved, _
                          oldXlScreenUpdating, oldXlScreenUpdatingSaved, _
                          oldXlEnableEvents, oldXlEnableEventsSaved, _
                          oldXlDisplayAlerts, oldXlDisplayAlertsSaved

    ResetViewFilter

    If Len(prevTable) > 0 Then Application.TableApply prevTable
    If Len(prevView) > 0 Then Application.ViewApply prevView

    If projectScreenSaved Then Application.ScreenUpdating = oldProjectScreenUpdating
    If projectCalcSaved Then Application.Calculation = originalProjectCalculation

    Application.StatusBar = False

    On Error GoTo 0

    MsgBox "Ошибка экспорта: " & errNum & " — " & errDesc, vbExclamation
End Sub

Private Sub SaveAndOptimizeExcelState(ByVal xlApp As Object, _
                                      ByRef oldCalc As Variant, _
                                      ByRef calcSaved As Boolean, _
                                      ByRef oldScreenUpdating As Boolean, _
                                      ByRef screenSaved As Boolean, _
                                      ByRef oldEnableEvents As Boolean, _
                                      ByRef eventsSaved As Boolean, _
                                      ByRef oldDisplayAlerts As Boolean, _
                                      ByRef alertsSaved As Boolean)
    On Error Resume Next

    Const xlCalculationManual As Long = -4135

    calcSaved = False
    screenSaved = False
    eventsSaved = False
    alertsSaved = False

    If xlApp Is Nothing Then Exit Sub

    Err.Clear
    oldCalc = xlApp.Calculation
    If Err.Number = 0 Then
        calcSaved = True
        Err.Clear
        xlApp.Calculation = xlCalculationManual
    End If
    Err.Clear

    oldScreenUpdating = xlApp.ScreenUpdating
    If Err.Number = 0 Then
        screenSaved = True
        Err.Clear
        xlApp.ScreenUpdating = False
    End If
    Err.Clear

    oldEnableEvents = xlApp.EnableEvents
    If Err.Number = 0 Then
        eventsSaved = True
        Err.Clear
        xlApp.EnableEvents = False
    End If
    Err.Clear

    oldDisplayAlerts = xlApp.DisplayAlerts
    If Err.Number = 0 Then
        alertsSaved = True
        Err.Clear
        xlApp.DisplayAlerts = False
    End If
    Err.Clear

    On Error GoTo 0
End Sub

Private Sub RestoreExcelStateSafe(ByVal xlApp As Object, _
                                  ByVal oldCalc As Variant, _
                                  ByVal calcSaved As Boolean, _
                                  ByVal oldScreenUpdating As Boolean, _
                                  ByVal screenSaved As Boolean, _
                                  ByVal oldEnableEvents As Boolean, _
                                  ByVal eventsSaved As Boolean, _
                                  ByVal oldDisplayAlerts As Boolean, _
                                  ByVal alertsSaved As Boolean)
    On Error Resume Next

    If xlApp Is Nothing Then Exit Sub

    If calcSaved Then
        Err.Clear
        xlApp.Calculation = oldCalc
        Err.Clear
    End If

    If screenSaved Then
        Err.Clear
        xlApp.ScreenUpdating = oldScreenUpdating
        Err.Clear
    End If

    If eventsSaved Then
        Err.Clear
        xlApp.EnableEvents = oldEnableEvents
        Err.Clear
    End If

    If alertsSaved Then
        Err.Clear
        xlApp.DisplayAlerts = oldDisplayAlerts
        Err.Clear
    End If

    On Error GoTo 0
End Sub

Private Function ExportPeriodByMonthsToWorkbook(ByVal wb As Object, _
                                                ByVal xlApp As Object, _
                                                ByVal SD0 As Date, _
                                                ByVal ED0 As Date, _
                                                ByRef monthData() As TMonthSummary, _
                                                ByVal totalMonths As Long) As Long
    On Error GoTo EH

    Dim shPeriod As Object
    Dim shMonth As Object

    Dim firstMonth As Date
    Dim lastMonth As Date
    Dim m As Date

    Dim SDm As Date
    Dim EDm As Date
    Dim monthEnd As Date

    Dim madeMonthSheets As Long

    Application.StatusBar = "Расчёт всего выбранного периода: " & _
                            Format(SD0, "dd.mm.yyyy") & " — " & Format(ED0, "dd.mm.yyyy")

    RunCalcAndFilter SD0, ED0, False, False

    Set shPeriod = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    shPeriod.Name = GetUniqueSheetName(wb, GetPeriodSheetName(SD0, ED0))

    ExportProjectPeriodToExcelSheet shPeriod, SD0, ED0

    firstMonth = DateSerial(Year(SD0), Month(SD0), 1)
    lastMonth = DateSerial(Year(ED0), Month(ED0), 1)

    m = firstMonth
    madeMonthSheets = 0

    Do While m <= lastMonth
        monthEnd = DateSerial(Year(m), Month(m) + 1, 0)

        SDm = MaxDate(SD0, m)
        EDm = MinDate(ED0, monthEnd)

        madeMonthSheets = madeMonthSheets + 1

        Application.StatusBar = "Расчёт месяца: " & Format(m, "mm.yyyy") & _
                                " | " & Format(SDm, "dd.mm.yyyy") & _
                                " — " & Format(EDm, "dd.mm.yyyy")

        RunCalcAndFilter SDm, EDm, False, False

        CollectSummaryForCurrentPeriod monthData(madeMonthSheets), SDm, EDm

        Set shMonth = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        shMonth.Name = GetUniqueSheetName(wb, Format(m, "mm.yyyy"))

        ExportProjectPeriodToExcelSheet shMonth, SDm, EDm

        m = DateAdd("m", 1, m)
    Loop

    ExportPeriodByMonthsToWorkbook = madeMonthSheets
    Exit Function

EH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetPeriodSheetName(ByVal SD0 As Date, ByVal ED0 As Date) As String
    GetPeriodSheetName = "Период " & Format(SD0, "dd.mm") & "-" & Format(ED0, "dd.mm.yy")
End Function

Private Sub ExportProjectPeriodToExcelSheet(ByVal sh As Object, ByVal SD0 As Date, ByVal ED0 As Date)
    On Error GoTo EH

    Dim headers As Variant
    headers = GetExportHeadersArray()

    sh.Cells.Clear
    sh.Range("A1").Resize(1, 33).Value = headers

    Dim includeSet As Object
    Set includeSet = BuildExportIdSetForPeriod(SD0, ED0)

    Dim rowCount As Long
    rowCount = CountExportRowsForPeriod(includeSet)

    If rowCount = 0 Then
        sh.Range("A2").Value = "Нет работ за период " & Format(SD0, "dd.mm.yyyy") & " — " & Format(ED0, "dd.mm.yyyy")
        FormatExportSheet sh, 1
        Exit Sub
    End If

    Dim arr() As Variant
    Dim rowOutlineLevels() As Long
    Dim rowSummaryFlags() As Boolean

    ReDim arr(1 To rowCount, 1 To 33)
    ReDim rowOutlineLevels(1 To rowCount)
    ReDim rowSummaryFlags(1 To rowCount)

    Dim t As Task
    Dim r As Long
    Dim key As String

    r = 0

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            key = CStr(t.UniqueID)

            If includeSet.Exists(key) Then
                r = r + 1

                rowOutlineLevels(r) = t.OutlineLevel
                rowSummaryFlags(r) = t.Summary

                arr(r, 1) = t.ID
                arr(r, 2) = t.Text2
                arr(r, 3) = t.Text7
                arr(r, 4) = t.OutlineLevel
                arr(r, 5) = t.Text4
                arr(r, 6) = t.Text10
                arr(r, 7) = t.Text20
                arr(r, 8) = t.Text21
                arr(r, 9) = t.Text22
                arr(r, 10) = t.Name
                arr(r, 11) = DurationMinutesToDays(t.Duration)
                arr(r, 12) = SafeDateValue(t.Start)
                arr(r, 13) = SafeDateValue(t.Finish)
                arr(r, 14) = t.Text1
                arr(r, 15) = ToD(t.Number1)
                arr(r, 16) = ToD(t.Number2)
                arr(r, 17) = ToD(t.Number3)
                arr(r, 18) = ToD(t.Number11)
                arr(r, 19) = ToD(t.Cost5)
                arr(r, 20) = ToD(t.Cost1)
                arr(r, 21) = ToD(t.Cost4)
                arr(r, 22) = ToD(t.Cost8)
                arr(r, 23) = ToD(t.Cost7)
                arr(r, 24) = t.ResourceNames
                arr(r, 25) = ToD(t.Number17)
                arr(r, 26) = ToD(t.Number6)
                arr(r, 27) = ToD(t.Number8)
                arr(r, 28) = ToD(t.Number18)
                arr(r, 29) = ToD(t.Number10)
                arr(r, 30) = ToD(t.Number12)
                arr(r, 31) = ToD(t.Cost2)
                arr(r, 32) = ToD(t.Cost3)
                arr(r, 33) = ToD(t.Cost6)
            End If
        End If
    Next t

    sh.Range("A2").Resize(rowCount, 33).Value = arr

    FormatExportSheet sh, rowCount + 1, rowOutlineLevels, rowSummaryFlags

    Exit Sub

EH:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function BuildExportIdSetForPeriod(ByVal SD0 As Date, ByVal ED0 As Date) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim t As Task
    Dim p As Task

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If Not t.Summary Then
                If TaskIntersectsPeriod(t, SD0, ED0) Then

                    If Not dict.Exists(CStr(t.UniqueID)) Then
                        dict.Add CStr(t.UniqueID), True
                    End If

                    Set p = t.OutlineParent

                    Do While Not p Is Nothing
                        If Not dict.Exists(CStr(p.UniqueID)) Then
                            dict.Add CStr(p.UniqueID), True
                        End If

                        Set p = p.OutlineParent
                    Loop
                End If
            End If
        End If
    Next t

    Set BuildExportIdSetForPeriod = dict
End Function

Private Function CountExportRowsForPeriod(ByVal includeSet As Object) As Long
    Dim t As Task
    Dim cnt As Long

    cnt = 0

    If includeSet Is Nothing Then
        CountExportRowsForPeriod = 0
        Exit Function
    End If

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If includeSet.Exists(CStr(t.UniqueID)) Then
                cnt = cnt + 1
            End If
        End If
    Next t

    CountExportRowsForPeriod = cnt
End Function

Private Function GetExportHeadersArray() As Variant
    Dim arr(1 To 1, 1 To 33) As Variant

    arr(1, 1) = "Ид."
    arr(1, 2) = "№ п/п"
    arr(1, 3) = "% Освоения ДС"
    arr(1, 4) = "Уровень структуры"
    arr(1, 5) = "Тип задачи"
    arr(1, 6) = "Выполнение"
    arr(1, 7) = "Объект"
    arr(1, 8) = "Подобъект"
    arr(1, 9) = "Раздел"
    arr(1, 10) = "Название задачи"
    arr(1, 11) = "Длительность, дн."
    arr(1, 12) = "Начало"
    arr(1, 13) = "Окончание"
    arr(1, 14) = "Ед. изм."
    arr(1, 15) = "Объем проект"
    arr(1, 16) = "Объем план"
    arr(1, 17) = "Объем факт"
    arr(1, 18) = "Объем факт (нак.)"
    arr(1, 19) = "Стоимость См (руб. с НДС)"
    arr(1, 20) = "Стоимость См план (с превыш.)"
    arr(1, 21) = "Стоимость См факт (руб. с НДС)"
    arr(1, 22) = "Стоимость См план (руб. с НДС)"
    arr(1, 23) = "Стоимость См факт (без превыш.)"
    arr(1, 24) = "Название ресурсов"
    arr(1, 25) = "ТЗр Смета план (ч/ч)"
    arr(1, 26) = "ТЗр Бюджет план (ч/ч)"
    arr(1, 27) = "ТЗр факт (ч/ч)"
    arr(1, 28) = "МиМ Смета план (м/ч)"
    arr(1, 29) = "МиМ Бюджет план (м/ч)"
    arr(1, 30) = "МиМ факт (м/ч)"
    arr(1, 31) = "Затраты персонал"
    arr(1, 32) = "Затраты МиМ"
    arr(1, 33) = "Затраты персонал+МиМ"

    GetExportHeadersArray = arr
End Function

Private Sub FormatExportSheet(ByVal sh As Object, _
                              ByVal lastRow As Long, _
                              Optional ByRef rowOutlineLevels As Variant, _
                              Optional ByRef rowSummaryFlags As Variant)
    On Error Resume Next

    Dim c As Long
    Dim r As Long
    Dim dataIndex As Long
    Dim lvl As Long
    Dim isSummary As Boolean

    With sh
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 10
        .Cells.WrapText = False

        .Range("A1:AG1").Font.Bold = True
        .Range("A1:AG1").Interior.Color = RGB(31, 78, 121)
        .Range("A1:AG1").Font.Color = RGB(255, 255, 255)
        .Range("A1:AG1").HorizontalAlignment = -4108
        .Range("A1:AG1").VerticalAlignment = -4108
        .Range("A1:AG1").WrapText = True

        If lastRow >= 1 Then
            .Range("A1:AG" & lastRow).Borders.LineStyle = 1
            .Range("A1:AG" & lastRow).Borders.Color = RGB(210, 210, 210)
        End If

        .Columns("K").NumberFormat = "0.00"
        .Columns("L:M").NumberFormat = "dd.mm.yyyy"
        .Columns("O:W").NumberFormat = "#,##0.00"
        .Columns("Y:AG").NumberFormat = "#,##0.00"

        If ExportColWidthsInitialized Then
            For c = 0 To 32
                .Columns(c + 1).ColumnWidth = ExportColWidths(c)
            Next c
        Else
            .Columns("A:AG").EntireColumn.AutoFit
        End If

        .Rows(1).RowHeight = 42

        If lastRow >= 2 Then
            If IsArray(rowOutlineLevels) And IsArray(rowSummaryFlags) Then
                For r = 2 To lastRow
                    dataIndex = r - 1

                    lvl = 1
                    isSummary = False

                    lvl = CLng(rowOutlineLevels(dataIndex))
                    isSummary = CBool(rowSummaryFlags(dataIndex))

                    If lvl < 1 Then lvl = 1

                    .Cells(r, 10).IndentLevel = MinLong(lvl - 1, 15)

                    If isSummary Then
                        .Rows(r).Font.Bold = True

                        Select Case lvl
                            Case 1
                                .Range("A" & r & ":AG" & r).Interior.Color = RGB(217, 225, 242)
                            Case 2
                                .Range("A" & r & ":AG" & r).Interior.Color = RGB(226, 239, 218)
                            Case 3
                                .Range("A" & r & ":AG" & r).Interior.Color = RGB(242, 242, 242)
                            Case Else
                                .Range("A" & r & ":AG" & r).Interior.Color = RGB(248, 248, 248)
                        End Select
                    Else
                        .Rows(r).Font.Bold = False
                        .Range("A" & r & ":AG" & r).Interior.Color = RGB(255, 255, 255)
                    End If
                Next r

                ApplyExcelProjectOutline sh, rowOutlineLevels, rowSummaryFlags, lastRow - 1
                .Outline.ShowLevels RowLevels:=4
            End If
        End If

        .Range("A1:AG" & lastRow).AutoFilter
    End With

    On Error GoTo 0
End Sub

Private Sub ApplyExcelProjectOutline(ByVal sh As Object, _
                                     ByRef rowOutlineLevels As Variant, _
                                     ByRef rowSummaryFlags As Variant, _
                                     ByVal rowCount As Long)
    On Error Resume Next

    If rowCount <= 0 Then Exit Sub
    If Not IsArray(rowOutlineLevels) Then Exit Sub
    If Not IsArray(rowSummaryFlags) Then Exit Sub

    Dim maxLevel As Long
    Dim i As Long
    Dim levelNum As Long
    Dim parentLevel As Long
    Dim endIndex As Long
    Dim childStartExcelRow As Long
    Dim childEndExcelRow As Long

    maxLevel = 1

    For i = 1 To rowCount
        If CLng(rowOutlineLevels(i)) > maxLevel Then maxLevel = CLng(rowOutlineLevels(i))
    Next i

    sh.Outline.SummaryRow = 0

    For levelNum = maxLevel To 1 Step -1
        For i = 1 To rowCount
            If CBool(rowSummaryFlags(i)) Then
                If CLng(rowOutlineLevels(i)) = levelNum Then

                    parentLevel = CLng(rowOutlineLevels(i))
                    endIndex = i

                    Do While endIndex + 1 <= rowCount
                        If CLng(rowOutlineLevels(endIndex + 1)) <= parentLevel Then Exit Do
                        endIndex = endIndex + 1
                    Loop

                    If endIndex > i Then
                        childStartExcelRow = i + 2
                        childEndExcelRow = endIndex + 1

                        If childEndExcelRow >= childStartExcelRow Then
                            sh.Rows(CStr(childStartExcelRow) & ":" & CStr(childEndExcelRow)).Group
                        End If
                    End If
                End If
            End If
        Next i
    Next levelNum

    On Error GoTo 0
End Sub

Private Function DurationMinutesToDays(ByVal durationValue As Variant) As Double
    On Error GoTo SafeExit

    DurationMinutesToDays = CDbl(durationValue) / 480#
    Exit Function

SafeExit:
    DurationMinutesToDays = 0
End Function

Private Function SafeDateValue(ByVal v As Variant) As Variant
    On Error GoTo SafeExit

    If IsDate(v) Then
        SafeDateValue = CDate(v)
    Else
        SafeDateValue = vbNullString
    End If

    Exit Function

SafeExit:
    SafeDateValue = vbNullString
End Function

' ============================================================
' СБОР СВОДНЫХ ДАННЫХ
' ============================================================

Private Sub CollectSummaryForCurrentPeriod(ByRef ms As TMonthSummary, _
                                           ByVal SD0 As Date, _
                                           ByVal ED0 As Date)
    On Error GoTo SafeExit

    ClearMonthSummary ms

    ms.StartDate = SD0
    ms.EndDate = ED0
    ms.Label = Format$(SD0, "mmmm yyyy")
    ms.YearNum = Year(SD0)
    ms.DaysInPeriod = CLng(Int(ED0) - Int(SD0) + 1)

    Dim t As Task
    Dim isEquip As Boolean
    Dim isSub As Boolean

    Dim planCost As Double
    Dim factCost As Double

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If Not t.Summary Then
                If t.OutlineLevel = 4 Then
                    If TaskIntersectsPeriod(t, SD0, ED0) Then

                        isEquip = IsEquipmentTask(CStr(t.Text4))
                        isSub = IsSubcontractTask(CStr(t.Text10))

                        planCost = ToD(t.Cost1)
                        factCost = ToD(t.Cost4)

                        ms.PlanTotal = ms.PlanTotal + planCost
                        ms.FactTotal = ms.FactTotal + factCost

                        If isEquip Then
                            ms.PlanEquip = ms.PlanEquip + planCost
                            ms.FactEquip = ms.FactEquip + factCost
                        Else
                            ms.PlanCMR = ms.PlanCMR + planCost
                            ms.FactCMR = ms.FactCMR + factCost
                        End If

                        ms.TzrSmeta = ms.TzrSmeta + ToD(t.Number17)

                        If isSub Then
                            ms.TzrSmetaSub = ms.TzrSmetaSub + ToD(t.Number17)
                        Else
                            ms.TzrBudgetSS = ms.TzrBudgetSS + ToD(t.Number6)
                            ms.TzrFactSS = ms.TzrFactSS + ToD(t.Number8)
                        End If

                        ms.MimSmeta = ms.MimSmeta + ToD(t.Number18)
                        ms.MimBudget = ms.MimBudget + ToD(t.Number10)
                        ms.MimFact = ms.MimFact + ToD(t.Number12)
                    End If
                End If
            End If
        End If
    Next t

    If ms.DaysInPeriod > 0 Then
        ms.PeopleSub = ms.TzrSmetaSub / ms.DaysInPeriod / 8#
        ms.PeopleSSPlanTzr = ms.TzrBudgetSS / ms.DaysInPeriod / 8#
    End If

    If ms.PlanTotal > 0 Then
        ms.PercentCompleteMoney = ms.FactTotal / ms.PlanTotal
    Else
        ms.PercentCompleteMoney = 0
    End If

SafeExit:
    On Error GoTo 0
End Sub

Private Sub ClearMonthSummary(ByRef ms As TMonthSummary)
    Dim z As TMonthSummary
    ms = z
End Sub

Private Function IsEquipmentTask(ByVal typeText As String) As Boolean
    Dim s As String

    s = UCase$(Trim$(typeText))
    s = Replace$(s, "|", "")
    s = Replace$(s, "Ё", "Е")

    IsEquipmentTask = (InStr(1, s, "ОБОРУДОВАНИЕ", vbTextCompare) > 0)
End Function

Private Function IsSubcontractTask(ByVal execText As String) As Boolean
    Dim s As String

    s = UCase$(Trim$(execText))
    s = Replace$(s, "Ё", "Е")

    IsSubcontractTask = (InStr(1, s, "СУБПОДРЯД", vbTextCompare) > 0)
End Function

Private Function ToD(ByVal v As Variant) As Double
    On Error GoTo SafeExit

    Dim s As String

    If IsNumeric(v) Then
        ToD = CDbl(v)
        Exit Function
    End If

    s = CStr(v)
    s = Replace$(s, ChrW(160), " ")
    s = Replace$(s, " ", "")
    s = Replace$(s, "руб.", "", , , vbTextCompare)
    s = Replace$(s, "руб", "", , , vbTextCompare)
    s = Replace$(s, "р.", "", , , vbTextCompare)
    s = Replace$(s, "р", "", , , vbTextCompare)
    s = Replace$(s, ChrW(&H20BD), "")
    s = Replace$(s, ",", ".")
    s = Trim$(s)

    If Len(s) = 0 Then
        ToD = 0
    Else
        ToD = Val(s)
    End If

    Exit Function

SafeExit:
    ToD = 0
End Function

' ============================================================
' СВОДНЫЙ ЛИСТ
' ============================================================

Private Sub BuildSummarySheet(ByVal ws As Object, _
                              ByVal xlApp As Object, _
                              ByRef monthData() As TMonthSummary, _
                              ByVal monthCount As Long, _
                              ByVal SD0 As Date, _
                              ByVal ED0 As Date)
    On Error GoTo EH

    If monthCount <= 0 Then Exit Sub

    ws.Cells.Clear

    Dim r As Long
    Dim i As Long
    Dim currentYear As Long

    Dim firstDataRow As Long
    Dim lastDataRow As Long

    Dim yearStartRow As Long
    Dim yearEndRow As Long
    Dim grandEndRow As Long

    Dim yearGroups As Collection
    Set yearGroups = New Collection

    firstDataRow = 5
    r = firstDataRow

    PrepareSummaryHeader ws, SD0, ED0

    currentYear = monthData(1).YearNum
    yearStartRow = r

    For i = 1 To monthCount

        If i > 1 Then
            If monthData(i).YearNum <> currentYear Then
                yearEndRow = r - 1

                WriteTotalFormulaRow ws, r, "Итог " & CStr(currentYear), yearStartRow, yearEndRow

                If yearEndRow >= yearStartRow Then
                    yearGroups.Add CStr(yearStartRow) & "|" & CStr(yearEndRow)
                End If

                r = r + 1

                currentYear = monthData(i).YearNum
                yearStartRow = r
            End If
        End If

        WriteSummaryRow ws, r, monthData(i).Label, monthData(i), True, False

        r = r + 1
    Next i

    yearEndRow = r - 1
    WriteTotalFormulaRow ws, r, "Итог " & CStr(currentYear), yearStartRow, yearEndRow

    If yearEndRow >= yearStartRow Then
        yearGroups.Add CStr(yearStartRow) & "|" & CStr(yearEndRow)
    End If

    r = r + 1

    grandEndRow = r - 2
    WriteGrandTotalFormulaRow ws, r, "ОБЩИЙ ИТОГ", firstDataRow, grandEndRow

    lastDataRow = r

    FormatSummarySheet ws, xlApp, firstDataRow, lastDataRow

    ApplyYearRowGrouping ws, yearGroups

    CreateSummaryCharts ws, monthData, monthCount, firstDataRow, lastDataRow

    Exit Sub

EH:
    MsgBox "Сводный лист сформирован с ошибкой: " & Err.Number & " — " & Err.Description, vbExclamation
End Sub

Private Sub ApplyYearRowGrouping(ByVal ws As Object, ByVal yearGroups As Collection)
    On Error Resume Next

    If yearGroups Is Nothing Then Exit Sub
    If yearGroups.Count = 0 Then Exit Sub

    Dim i As Long
    Dim p() As String
    Dim startRow As Long
    Dim endRow As Long

    ws.Outline.SummaryRow = 1

    For i = 1 To yearGroups.Count
        p = Split(CStr(yearGroups(i)), "|")

        If UBound(p) = 1 Then
            startRow = CLng(p(0))
            endRow = CLng(p(1))

            If endRow > startRow Then
                ws.Rows(CStr(startRow) & ":" & CStr(endRow)).Group
            End If
        End If
    Next i

    ws.Outline.ShowLevels RowLevels:=2

    On Error GoTo 0
End Sub

Private Sub PrepareSummaryHeader(ByVal ws As Object, ByVal SD0 As Date, ByVal ED0 As Date)
    With ws
        .Range("A1:AA1").Merge
        .Range("A1").Value = GetProjectSummaryTitle()
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .Range("A1").HorizontalAlignment = -4108
        .Range("A1").VerticalAlignment = -4108
        .Range("A1").Interior.Color = RGB(31, 78, 121)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        .Range("A2:B3").Merge
        .Range("A2").Value = vbNullString

        .Range("A4").Value = "Период"
        .Range("B4").Value = "% освоения ДС"

        .Range("C2:K2").Merge
        .Range("C2").Value = "Сметная стоимость, руб. с НДС"

        .Range("C3:E3").Merge
        .Range("C3").Value = "Базовый план"

        .Range("F3:H3").Merge
        .Range("F3").Value = "План"

        .Range("I3:K3").Merge
        .Range("I3").Value = "Факт"

        .Range("L2:Q2").Merge
        .Range("L2").Value = "Трудозатраты, ч/ч"

        .Range("R2:T2").Merge
        .Range("R2").Value = "Машины и механизмы, м/ч"

        .Range("U2:AA2").Merge
        .Range("U2").Value = "Технический блок для диаграмм"

        .Range("C4").Value = "Сметная стоимость"
        .Range("D4").Value = "в т.ч. СМР"
        .Range("E4").Value = "в т.ч. Оборудование"

        .Range("F4").Value = "Сметная стоимость"
        .Range("G4").Value = "в т.ч. СМР"
        .Range("H4").Value = "в т.ч. Оборудование"

        .Range("I4").Value = "Сметная стоимость"
        .Range("J4").Value = "в т.ч. СМР"
        .Range("K4").Value = "в т.ч. Оборудование"

        .Range("L4").Value = "ТЗр Смета"
        .Range("M4").Value = "ТЗр Смета субподряд"
        .Range("N4").Value = "ТЗр Бюджет СС"
        .Range("O4").Value = "Кол-во чел. субподряд"
        .Range("P4").Value = "Кол-во чел. СС по плану ТЗр"
        .Range("Q4").Value = "ТЗр факт СС"

        .Range("R4").Value = "МиМ Смета план"
        .Range("S4").Value = "МиМ Бюджет план"
        .Range("T4").Value = "МиМ факт"

        .Range("U4").Value = "План накопительный"
        .Range("V4").Value = "Факт накопительный"
        .Range("W4").Value = "БП накопительный"
        .Range("X4").Value = "% исполнительный"
        .Range("Y4").Value = "Кол-во дней в месяце"
        .Range("Z4").Value = "Кол-во чел. субподряд"
        .Range("AA4").Value = "Кол-во человек СС по бюджету"
    End With
End Sub

Private Function GetProjectSummaryTitle() As String
    On Error GoTo SafeExit

    If Not ActiveProject.ProjectSummaryTask Is Nothing Then
        If Len(Trim$(ActiveProject.ProjectSummaryTask.Name)) > 0 Then
            GetProjectSummaryTitle = ActiveProject.ProjectSummaryTask.Name
            Exit Function
        End If
    End If

SafeExit:
    If Len(Trim$(GetProjectSummaryTitle)) = 0 Then
        GetProjectSummaryTitle = ActiveProject.Name
    End If
End Function

Private Function GetMonthOnlyCumulativeFormula(ByVal sourceColumn As String, ByVal rowNum As Long) As String
    GetMonthOnlyCumulativeFormula = _
        "=SUMIFS($" & sourceColumn & "$5:" & sourceColumn & rowNum & _
        ",$A$5:A" & rowNum & ",""<>Итог*""" & _
        ",$A$5:A" & rowNum & ",""<>ОБЩИЙ*"")"
End Function

Private Sub WriteSummaryRow(ByVal ws As Object, _
                            ByVal rowNum As Long, _
                            ByVal rowLabel As String, _
                            ByRef s As TMonthSummary, _
                            ByVal isMonthRow As Boolean, _
                            ByVal isTotalRow As Boolean)
    With ws
        .Cells(rowNum, 1).Value = rowLabel

        If s.PlanTotal > 0 Then
            .Cells(rowNum, 2).Value = s.FactTotal / s.PlanTotal
        Else
            .Cells(rowNum, 2).Value = 0
        End If

        .Cells(rowNum, 3).Value = vbNullString
        .Cells(rowNum, 4).Value = vbNullString
        .Cells(rowNum, 5).Value = vbNullString

        .Cells(rowNum, 6).Value = Round(s.PlanTotal, 2)
        .Cells(rowNum, 7).Value = Round(s.PlanCMR, 2)
        .Cells(rowNum, 8).Value = Round(s.PlanEquip, 2)

        .Cells(rowNum, 9).Value = Round(s.FactTotal, 2)
        .Cells(rowNum, 10).Value = Round(s.FactCMR, 2)
        .Cells(rowNum, 11).Value = Round(s.FactEquip, 2)

        .Cells(rowNum, 12).Value = Round(s.TzrSmeta, 2)
        .Cells(rowNum, 13).Value = Round(s.TzrSmetaSub, 2)
        .Cells(rowNum, 14).Value = Round(s.TzrBudgetSS, 2)

        If s.DaysInPeriod > 0 Then
            .Cells(rowNum, 15).Value = Round(s.TzrSmetaSub / s.DaysInPeriod / 8#, 0)
            .Cells(rowNum, 16).Value = Round(s.TzrBudgetSS / s.DaysInPeriod / 8#, 0)
        Else
            .Cells(rowNum, 15).Value = 0
            .Cells(rowNum, 16).Value = 0
        End If

        .Cells(rowNum, 17).Value = Round(s.TzrFactSS, 2)

        .Cells(rowNum, 18).Value = Round(s.MimSmeta, 2)
        .Cells(rowNum, 19).Value = Round(s.MimBudget, 2)
        .Cells(rowNum, 20).Value = Round(s.MimFact, 2)

        If isMonthRow Then
            .Cells(rowNum, 21).Formula = GetMonthOnlyCumulativeFormula("F", rowNum)
            .Cells(rowNum, 22).Formula = GetMonthOnlyCumulativeFormula("I", rowNum)
            .Cells(rowNum, 23).Formula = GetMonthOnlyCumulativeFormula("C", rowNum)
            .Cells(rowNum, 24).FormulaR1C1 = "=IF(RC21=0,0,RC22/RC21)"

            .Cells(rowNum, 25).Value = s.DaysInPeriod
            .Cells(rowNum, 26).Value = .Cells(rowNum, 15).Value
            .Cells(rowNum, 27).Value = .Cells(rowNum, 16).Value
        Else
            .Cells(rowNum, 21).Value = vbNullString
            .Cells(rowNum, 22).Value = vbNullString
            .Cells(rowNum, 23).Value = vbNullString
            .Cells(rowNum, 24).Value = vbNullString
            .Cells(rowNum, 25).Value = s.DaysInPeriod

            If s.DaysInPeriod > 0 Then
                .Cells(rowNum, 26).Value = Round(s.TzrSmetaSub / s.DaysInPeriod / 8#, 0)
                .Cells(rowNum, 27).Value = Round(s.TzrBudgetSS / s.DaysInPeriod / 8#, 0)
            Else
                .Cells(rowNum, 26).Value = 0
                .Cells(rowNum, 27).Value = 0
            End If
        End If

        If isTotalRow Then
            .Range(.Cells(rowNum, 1), .Cells(rowNum, 27)).Font.Bold = True
            .Range(.Cells(rowNum, 1), .Cells(rowNum, 27)).Interior.Color = RGB(217, 217, 217)
        End If
    End With
End Sub

Private Sub WriteTotalFormulaRow(ByVal ws As Object, _
                                 ByVal rowNum As Long, _
                                 ByVal rowLabel As String, _
                                 ByVal startRow As Long, _
                                 ByVal endRow As Long)
    With ws
        .Cells(rowNum, 1).Value = rowLabel

        .Cells(rowNum, 2).FormulaR1C1 = "=IF(RC6=0,0,RC9/RC6)"

        .Cells(rowNum, 3).Formula = "=SUM(C" & startRow & ":C" & endRow & ")"
        .Cells(rowNum, 4).Formula = "=SUM(D" & startRow & ":D" & endRow & ")"
        .Cells(rowNum, 5).Formula = "=SUM(E" & startRow & ":E" & endRow & ")"

        .Cells(rowNum, 6).Formula = "=SUM(F" & startRow & ":F" & endRow & ")"
        .Cells(rowNum, 7).Formula = "=SUM(G" & startRow & ":G" & endRow & ")"
        .Cells(rowNum, 8).Formula = "=SUM(H" & startRow & ":H" & endRow & ")"

        .Cells(rowNum, 9).Formula = "=SUM(I" & startRow & ":I" & endRow & ")"
        .Cells(rowNum, 10).Formula = "=SUM(J" & startRow & ":J" & endRow & ")"
        .Cells(rowNum, 11).Formula = "=SUM(K" & startRow & ":K" & endRow & ")"

        .Cells(rowNum, 12).Formula = "=SUM(L" & startRow & ":L" & endRow & ")"
        .Cells(rowNum, 13).Formula = "=SUM(M" & startRow & ":M" & endRow & ")"
        .Cells(rowNum, 14).Formula = "=SUM(N" & startRow & ":N" & endRow & ")"

        .Cells(rowNum, 15).Formula = "=SUM(O" & startRow & ":O" & endRow & ")"
        .Cells(rowNum, 16).Formula = "=SUM(P" & startRow & ":P" & endRow & ")"
        .Cells(rowNum, 17).Formula = "=SUM(Q" & startRow & ":Q" & endRow & ")"

        .Cells(rowNum, 18).Formula = "=SUM(R" & startRow & ":R" & endRow & ")"
        .Cells(rowNum, 19).Formula = "=SUM(S" & startRow & ":S" & endRow & ")"
        .Cells(rowNum, 20).Formula = "=SUM(T" & startRow & ":T" & endRow & ")"

        .Cells(rowNum, 21).Formula = "=U" & endRow
        .Cells(rowNum, 22).Formula = "=V" & endRow
        .Cells(rowNum, 23).Formula = "=W" & endRow
        .Cells(rowNum, 24).FormulaR1C1 = "=IF(RC21=0,0,RC22/RC21)"

        .Cells(rowNum, 25).Formula = "=SUM(Y" & startRow & ":Y" & endRow & ")"
        .Cells(rowNum, 26).Formula = "=SUM(Z" & startRow & ":Z" & endRow & ")"
        .Cells(rowNum, 27).Formula = "=SUM(AA" & startRow & ":AA" & endRow & ")"

        .Range(.Cells(rowNum, 1), .Cells(rowNum, 27)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 27)).Interior.Color = RGB(217, 217, 217)
    End With
End Sub

Private Sub WriteGrandTotalFormulaRow(ByVal ws As Object, _
                                      ByVal rowNum As Long, _
                                      ByVal rowLabel As String, _
                                      ByVal startRow As Long, _
                                      ByVal endRow As Long)
    Dim r As Long
    Dim lastMonthRow As Long

    Dim formulaC As String, formulaD As String, formulaE As String
    Dim formulaF As String, formulaG As String, formulaH As String
    Dim formulaI As String, formulaJ As String, formulaK As String
    Dim formulaL As String, formulaM As String, formulaN As String
    Dim formulaO As String, formulaP As String, formulaQ As String
    Dim formulaR As String, formulaS As String, formulaT As String
    Dim formulaY As String, formulaZ As String, formulaAA As String

    lastMonthRow = 0

    For r = startRow To endRow
        If InStr(1, CStr(ws.Cells(r, 1).Value), "Итог", vbTextCompare) = 0 And _
           InStr(1, CStr(ws.Cells(r, 1).Value), "ОБЩИЙ", vbTextCompare) = 0 Then

            lastMonthRow = r

            AppendSumAddress formulaC, "C" & r
            AppendSumAddress formulaD, "D" & r
            AppendSumAddress formulaE, "E" & r

            AppendSumAddress formulaF, "F" & r
            AppendSumAddress formulaG, "G" & r
            AppendSumAddress formulaH, "H" & r

            AppendSumAddress formulaI, "I" & r
            AppendSumAddress formulaJ, "J" & r
            AppendSumAddress formulaK, "K" & r

            AppendSumAddress formulaL, "L" & r
            AppendSumAddress formulaM, "M" & r
            AppendSumAddress formulaN, "N" & r

            AppendSumAddress formulaO, "O" & r
            AppendSumAddress formulaP, "P" & r
            AppendSumAddress formulaQ, "Q" & r

            AppendSumAddress formulaR, "R" & r
            AppendSumAddress formulaS, "S" & r
            AppendSumAddress formulaT, "T" & r

            AppendSumAddress formulaY, "Y" & r
            AppendSumAddress formulaZ, "Z" & r
            AppendSumAddress formulaAA, "AA" & r
        End If
    Next r

    With ws
        .Cells(rowNum, 1).Value = rowLabel

        .Cells(rowNum, 2).FormulaR1C1 = "=IF(RC6=0,0,RC9/RC6)"

        .Cells(rowNum, 3).Formula = "=SUM(" & formulaC & ")"
        .Cells(rowNum, 4).Formula = "=SUM(" & formulaD & ")"
        .Cells(rowNum, 5).Formula = "=SUM(" & formulaE & ")"

        .Cells(rowNum, 6).Formula = "=SUM(" & formulaF & ")"
        .Cells(rowNum, 7).Formula = "=SUM(" & formulaG & ")"
        .Cells(rowNum, 8).Formula = "=SUM(" & formulaH & ")"

        .Cells(rowNum, 9).Formula = "=SUM(" & formulaI & ")"
        .Cells(rowNum, 10).Formula = "=SUM(" & formulaJ & ")"
        .Cells(rowNum, 11).Formula = "=SUM(" & formulaK & ")"

        .Cells(rowNum, 12).Formula = "=SUM(" & formulaL & ")"
        .Cells(rowNum, 13).Formula = "=SUM(" & formulaM & ")"
        .Cells(rowNum, 14).Formula = "=SUM(" & formulaN & ")"

        .Cells(rowNum, 15).Formula = "=SUM(" & formulaO & ")"
        .Cells(rowNum, 16).Formula = "=SUM(" & formulaP & ")"
        .Cells(rowNum, 17).Formula = "=SUM(" & formulaQ & ")"

        .Cells(rowNum, 18).Formula = "=SUM(" & formulaR & ")"
        .Cells(rowNum, 19).Formula = "=SUM(" & formulaS & ")"
        .Cells(rowNum, 20).Formula = "=SUM(" & formulaT & ")"

        If lastMonthRow > 0 Then
            .Cells(rowNum, 21).Formula = "=U" & lastMonthRow
            .Cells(rowNum, 22).Formula = "=V" & lastMonthRow
            .Cells(rowNum, 23).Formula = "=W" & lastMonthRow
            .Cells(rowNum, 24).FormulaR1C1 = "=IF(RC21=0,0,RC22/RC21)"
        Else
            .Cells(rowNum, 21).Value = vbNullString
            .Cells(rowNum, 22).Value = vbNullString
            .Cells(rowNum, 23).Value = vbNullString
            .Cells(rowNum, 24).Value = vbNullString
        End If

        .Cells(rowNum, 25).Formula = "=SUM(" & formulaY & ")"
        .Cells(rowNum, 26).Formula = "=SUM(" & formulaZ & ")"
        .Cells(rowNum, 27).Formula = "=SUM(" & formulaAA & ")"

        .Range(.Cells(rowNum, 1), .Cells(rowNum, 27)).Font.Bold = True
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 27)).Interior.Color = RGB(31, 78, 121)
        .Range(.Cells(rowNum, 1), .Cells(rowNum, 27)).Font.Color = RGB(255, 255, 255)
    End With
End Sub

Private Sub AppendSumAddress(ByRef formulaText As String, ByVal cellAddress As String)
    If Len(formulaText) = 0 Then
        formulaText = cellAddress
    Else
        formulaText = formulaText & "," & cellAddress
    End If
End Sub

Private Sub FormatSummarySheet(ByVal ws As Object, _
                               ByVal xlApp As Object, _
                               ByVal firstDataRow As Long, _
                               ByVal lastDataRow As Long)
    On Error Resume Next

    Dim r As Long

    With ws
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 12
        .Cells.WrapText = True

        .Range("A2:AA4").Font.Bold = True
        .Range("A2:AA4").HorizontalAlignment = -4108
        .Range("A2:AA4").VerticalAlignment = -4108

        .Range("A2:B4").Interior.Color = RGB(242, 242, 242)

        .Range("C2:E4").Interior.Color = RGB(217, 225, 242)
        .Range("F2:H4").Interior.Color = RGB(226, 239, 218)
        .Range("I2:K4").Interior.Color = RGB(221, 235, 247)

        .Range("L2:N4").Interior.Color = RGB(252, 228, 214)
        .Range("O2:O4").Interior.Color = RGB(244, 176, 132)
        .Range("P2:P4").Interior.Color = RGB(189, 215, 238)
        .Range("Q2:Q4").Interior.Color = RGB(252, 228, 214)

        .Range("R2:T4").Interior.Color = RGB(234, 241, 221)
        .Range("U2:AA4").Interior.Color = RGB(226, 239, 218)

        .Range("A" & firstDataRow & ":B" & lastDataRow).Interior.Color = RGB(250, 250, 250)
        .Range("C" & firstDataRow & ":E" & lastDataRow).Interior.Color = RGB(242, 246, 252)
        .Range("F" & firstDataRow & ":H" & lastDataRow).Interior.Color = RGB(244, 249, 239)
        .Range("I" & firstDataRow & ":K" & lastDataRow).Interior.Color = RGB(242, 248, 252)
        .Range("L" & firstDataRow & ":Q" & lastDataRow).Interior.Color = RGB(255, 248, 245)
        .Range("R" & firstDataRow & ":T" & lastDataRow).Interior.Color = RGB(248, 251, 243)
        .Range("U" & firstDataRow & ":AA" & lastDataRow).Interior.Color = RGB(241, 248, 239)

        .Range("A1:AA" & lastDataRow).Borders.LineStyle = 1
        .Range("A1:AA" & lastDataRow).Borders.Weight = 2
        .Range("A1:AA" & lastDataRow).Borders.Color = RGB(166, 166, 166)

        .Range("B" & firstDataRow & ":B" & lastDataRow).NumberFormat = "0.00%"
        .Range("C" & firstDataRow & ":N" & lastDataRow).NumberFormat = "#,##0.00"
        .Range("O" & firstDataRow & ":P" & lastDataRow).NumberFormat = "0"
        .Range("Q" & firstDataRow & ":T" & lastDataRow).NumberFormat = "#,##0.00"
        .Range("U" & firstDataRow & ":W" & lastDataRow).NumberFormat = "#,##0.00"
        .Range("X" & firstDataRow & ":X" & lastDataRow).NumberFormat = "0.00%"
        .Range("Y" & firstDataRow & ":AA" & lastDataRow).NumberFormat = "0"

        .Columns("A").ColumnWidth = 16
        .Columns("B").ColumnWidth = 13
        .Columns("C:K").ColumnWidth = 17
        .Columns("L:N").ColumnWidth = 15
        .Columns("O:P").ColumnWidth = 15
        .Columns("Q:T").ColumnWidth = 15
        .Columns("U:AA").ColumnWidth = 16

        .Rows("1:1").RowHeight = 34
        .Rows("2:3").RowHeight = 28
        .Rows("4:4").RowHeight = 56
        .Rows(firstDataRow & ":" & lastDataRow).RowHeight = 24

        For r = firstDataRow To lastDataRow
            If InStr(1, CStr(.Cells(r, 1).Value), "Итог", vbTextCompare) > 0 Then
                .Range(.Cells(r, 1), .Cells(r, 27)).Font.Bold = True
                .Range(.Cells(r, 1), .Cells(r, 27)).Interior.Color = RGB(217, 217, 217)
                .Range(.Cells(r, 1), .Cells(r, 27)).Font.Color = RGB(0, 0, 0)
            End If

            If InStr(1, CStr(.Cells(r, 1).Value), "ОБЩИЙ ИТОГ", vbTextCompare) > 0 Then
                .Range(.Cells(r, 1), .Cells(r, 27)).Font.Bold = True
                .Range(.Cells(r, 1), .Cells(r, 27)).Interior.Color = RGB(31, 78, 121)
                .Range(.Cells(r, 1), .Cells(r, 27)).Font.Color = RGB(255, 255, 255)
            End If
        Next r

        .Range("A4:AA" & lastDataRow).AutoFilter

        .Activate
        .Range("A5").Select
        xlApp.ActiveWindow.FreezePanes = True

        .PageSetup.PrintArea = "$A$1:$T$" & CStr(lastDataRow)
        .PageSetup.Orientation = 2
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False

        .Columns("U:AA").Group
        .Outline.ShowLevels ColumnLevels:=2
    End With

    On Error GoTo 0
End Sub

' ============================================================
' ДИАГРАММЫ
' ============================================================

Private Sub CreateSummaryCharts(ByVal ws As Object, _
                                ByRef monthData() As TMonthSummary, _
                                ByVal monthCount As Long, _
                                ByVal firstDataRow As Long, _
                                ByVal lastDataRow As Long)
    On Error GoTo SafeExit

    If monthCount <= 0 Then Exit Sub

    Dim chartSourceStartRow As Long
    Dim chartSourceLastRow As Long

    chartSourceStartRow = 2

    BuildChartSourceTable ws, monthData, monthCount, firstDataRow, lastDataRow, chartSourceStartRow, chartSourceLastRow

    Const xlColumnClustered As Long = 51
    Const xlLine As Long = 4
    Const xlLegendPositionBottom As Long = -4107

    Dim topPos As Double
    Dim leftPos As Double
    Dim tableWidth As Double
    Dim gap As Double
    Dim chartHeight As Double
    Dim chart1Width As Double
    Dim chart2Width As Double

    topPos = ws.Rows(lastDataRow + 3).Top
    leftPos = ws.Columns("A").Left
    tableWidth = ws.Range("A1:AA1").Width
    gap = 10
    chartHeight = 350

    chart1Width = (tableWidth - gap) * 0.545
    chart2Width = (tableWidth - gap) * 0.455

    Dim co1 As Object
    Dim ch1 As Object

    Set co1 = ws.ChartObjects.Add(leftPos, topPos, chart1Width, chartHeight)
    Set ch1 = co1.Chart

    ch1.ChartType = xlColumnClustered
    ch1.HasTitle = True
    ch1.ChartTitle.Text = "Освоение ДС"
    ch1.Legend.Position = xlLegendPositionBottom
    ch1.PlotVisibleOnly = False

    ClearChartSeries ch1

    With ch1.SeriesCollection.NewSeries
        .Name = "БП"
        .XValues = ws.Range("AC" & chartSourceStartRow & ":AC" & chartSourceLastRow)
        .Values = ws.Range("AD" & chartSourceStartRow & ":AD" & chartSourceLastRow)
        .ChartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(184, 204, 228)
        .Format.Line.ForeColor.RGB = RGB(184, 204, 228)
    End With

    With ch1.SeriesCollection.NewSeries
        .Name = "План"
        .XValues = ws.Range("AC" & chartSourceStartRow & ":AC" & chartSourceLastRow)
        .Values = ws.Range("AE" & chartSourceStartRow & ":AE" & chartSourceLastRow)
        .ChartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(155, 194, 230)
        .Format.Line.ForeColor.RGB = RGB(155, 194, 230)
    End With

    With ch1.SeriesCollection.NewSeries
        .Name = "Факт"
        .XValues = ws.Range("AC" & chartSourceStartRow & ":AC" & chartSourceLastRow)
        .Values = ws.Range("AF" & chartSourceStartRow & ":AF" & chartSourceLastRow)
        .ChartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(46, 117, 182)
        .Format.Line.ForeColor.RGB = RGB(46, 117, 182)
    End With

    With ch1.SeriesCollection.NewSeries
        .Name = "БП накопит."
        .XValues = ws.Range("AC" & chartSourceStartRow & ":AC" & chartSourceLastRow)
        .Values = ws.Range("AG" & chartSourceStartRow & ":AG" & chartSourceLastRow)
        .ChartType = xlLine
        .AxisGroup = 1
        .Format.Line.ForeColor.RGB = RGB(127, 127, 127)
        .Format.Line.Weight = 2.25
        .MarkerStyle = 8
        .MarkerSize = 4
        .MarkerForegroundColor = RGB(127, 127, 127)
        .MarkerBackgroundColor = RGB(127, 127, 127)
    End With

    With ch1.SeriesCollection.NewSeries
        .Name = "План накопит."
        .XValues = ws.Range("AC" & chartSourceStartRow & ":AC" & chartSourceLastRow)
        .Values = ws.Range("AH" & chartSourceStartRow & ":AH" & chartSourceLastRow)
        .ChartType = xlLine
        .AxisGroup = 1
        .Format.Line.ForeColor.RGB = RGB(91, 155, 213)
        .Format.Line.Weight = 2.75
        .MarkerStyle = 8
        .MarkerSize = 4
        .MarkerForegroundColor = RGB(91, 155, 213)
        .MarkerBackgroundColor = RGB(91, 155, 213)
    End With

    With ch1.SeriesCollection.NewSeries
        .Name = "Факт накопит."
        .XValues = ws.Range("AC" & chartSourceStartRow & ":AC" & chartSourceLastRow)
        .Values = ws.Range("AI" & chartSourceStartRow & ":AI" & chartSourceLastRow)
        .ChartType = xlLine
        .AxisGroup = 1
        .Format.Line.ForeColor.RGB = RGB(0, 84, 159)
        .Format.Line.Weight = 3
        .MarkerStyle = 8
        .MarkerSize = 4
        .MarkerForegroundColor = RGB(0, 84, 159)
        .MarkerBackgroundColor = RGB(0, 84, 159)
    End With

    FormatMoneyChart ch1

    Dim co2 As Object
    Dim ch2 As Object

    Set co2 = ws.ChartObjects.Add(leftPos + chart1Width + gap, topPos, chart2Width, chartHeight)
    Set ch2 = co2.Chart

    ch2.ChartType = xlColumnClustered
    ch2.HasTitle = True
    ch2.ChartTitle.Text = "Кол-во человек"
    ch2.Legend.Position = xlLegendPositionBottom
    ch2.PlotVisibleOnly = False

    ClearChartSeries ch2

    With ch2.SeriesCollection.NewSeries
        .Name = "Субподряд"
        .XValues = ws.Range("AC" & chartSourceStartRow & ":AC" & chartSourceLastRow)
        .Values = ws.Range("AJ" & chartSourceStartRow & ":AJ" & chartSourceLastRow)
        .ChartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        .Format.Line.ForeColor.RGB = RGB(237, 125, 49)
    End With

    With ch2.SeriesCollection.NewSeries
        .Name = "Собственные силы"
        .XValues = ws.Range("AC" & chartSourceStartRow & ":AC" & chartSourceLastRow)
        .Values = ws.Range("AK" & chartSourceStartRow & ":AK" & chartSourceLastRow)
        .ChartType = xlColumnClustered
        .Format.Fill.ForeColor.RGB = RGB(91, 155, 213)
        .Format.Line.ForeColor.RGB = RGB(91, 155, 213)
    End With

    FormatPeopleChart ch2

    ws.Columns("AC:AK").Hidden = True

SafeExit:
    On Error GoTo 0
End Sub

Private Sub BuildChartSourceTable(ByVal ws As Object, _
                                  ByRef monthData() As TMonthSummary, _
                                  ByVal monthCount As Long, _
                                  ByVal firstDataRow As Long, _
                                  ByVal lastDataRow As Long, _
                                  ByVal outStartRow As Long, _
                                  ByRef outLastRow As Long)
    On Error Resume Next

    Dim r As Long
    Dim outR As Long
    Dim idx As Long
    Dim labelText As String
    Dim lastFactIdx As Long

    lastFactIdx = GetLastFactMonthIndex(monthData, monthCount)

    ws.Range("AC:AK").Clear
    ws.Columns("AC").NumberFormat = "@"

    ws.Range("AC1").Value = "Период"
    ws.Range("AD1").Value = "БП"
    ws.Range("AE1").Value = "План"
    ws.Range("AF1").Value = "Факт"
    ws.Range("AG1").Value = "БП накопит."
    ws.Range("AH1").Value = "План накопит."
    ws.Range("AI1").Value = "Факт накопит."
    ws.Range("AJ1").Value = "Субподряд"
    ws.Range("AK1").Value = "Собственные силы"

    outR = outStartRow
    idx = 1

    For r = firstDataRow To lastDataRow
        labelText = CStr(ws.Cells(r, 1).Value)

        If Len(labelText) > 0 Then
            If InStr(1, labelText, "Итог", vbTextCompare) = 0 And _
               InStr(1, labelText, "ОБЩИЙ", vbTextCompare) = 0 Then

                If idx <= monthCount Then
                    ws.Cells(outR, "AC").NumberFormat = "@"
                    ws.Cells(outR, "AC").Value = Format$(monthData(idx).StartDate, "mm.yyyy")
                Else
                    ws.Cells(outR, "AC").NumberFormat = "@"
                    ws.Cells(outR, "AC").Value = CStr(ws.Cells(r, 1).Value)
                End If

                ws.Cells(outR, "AD").Formula = "=C" & r
                ws.Cells(outR, "AE").Formula = "=F" & r
                ws.Cells(outR, "AF").Formula = "=I" & r
                ws.Cells(outR, "AG").Formula = "=W" & r
                ws.Cells(outR, "AH").Formula = "=U" & r

                If idx <= lastFactIdx Then
                    ws.Cells(outR, "AI").Formula = "=V" & r
                Else
                    ws.Cells(outR, "AI").Formula = "=NA()"
                End If

                ws.Cells(outR, "AJ").Formula = "=O" & r
                ws.Cells(outR, "AK").Formula = "=P" & r

                outR = outR + 1
                idx = idx + 1
            End If
        End If
    Next r

    outLastRow = outR - 1

    ws.Range("AD" & outStartRow & ":AI" & outLastRow).NumberFormat = "#,##0"
    ws.Range("AJ" & outStartRow & ":AK" & outLastRow).NumberFormat = "0"

    On Error GoTo 0
End Sub

Private Function GetLastFactMonthIndex(ByRef monthData() As TMonthSummary, ByVal monthCount As Long) As Long
    Dim i As Long

    GetLastFactMonthIndex = 0

    For i = monthCount To 1 Step -1
        If Abs(monthData(i).FactTotal) > 0.000001 Then
            GetLastFactMonthIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub ClearChartSeries(ByVal ch As Object)
    On Error Resume Next

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    On Error GoTo 0
End Sub

Private Sub FormatMoneyChart(ByVal ch As Object)
    On Error Resume Next

    Const xlCategory As Long = 1
    Const xlValue As Long = 2
    Const xlPrimary As Long = 1

    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)

    ch.ChartArea.Format.Line.Visible = True
    ch.ChartArea.Format.Line.ForeColor.RGB = RGB(150, 150, 150)
    ch.ChartArea.Format.Line.Weight = 1.25

    ch.RoundedCorners = True

    ch.ChartTitle.Font.Name = "Calibri"
    ch.ChartTitle.Font.Size = 14
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = RGB(89, 89, 89)

    ch.Legend.Font.Name = "Calibri"
    ch.Legend.Font.Size = 9

    ch.Axes(xlCategory).TickLabels.Font.Name = "Calibri"
    ch.Axes(xlCategory).TickLabels.Font.Size = 8
    ch.Axes(xlCategory).TickLabels.Orientation = 45

    ch.Axes(xlValue, xlPrimary).TickLabels.Font.Name = "Calibri"
    ch.Axes(xlValue, xlPrimary).TickLabels.Font.Size = 9
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0 ₽"
    ch.Axes(xlValue, xlPrimary).MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
    ch.Axes(xlValue, xlPrimary).MajorGridlines.Format.Line.Weight = 0.75

    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(170, 170, 170)
    ch.Axes(xlValue, xlPrimary).Format.Line.ForeColor.RGB = RGB(170, 170, 170)

    ch.HasDataTable = False

    On Error GoTo 0
End Sub

Private Sub FormatPeopleChart(ByVal ch As Object)
    On Error Resume Next

    Const xlCategory As Long = 1
    Const xlValue As Long = 2
    Const xlPrimary As Long = 1

    ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)

    ch.ChartArea.Format.Line.Visible = True
    ch.ChartArea.Format.Line.ForeColor.RGB = RGB(150, 150, 150)
    ch.ChartArea.Format.Line.Weight = 1.25

    ch.RoundedCorners = True

    ch.ChartTitle.Font.Name = "Calibri"
    ch.ChartTitle.Font.Size = 14
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = RGB(89, 89, 89)

    ch.Legend.Font.Name = "Calibri"
    ch.Legend.Font.Size = 9

    ch.Axes(xlCategory).TickLabels.Font.Name = "Calibri"
    ch.Axes(xlCategory).TickLabels.Font.Size = 8
    ch.Axes(xlCategory).TickLabels.Orientation = 45

    ch.Axes(xlValue, xlPrimary).TickLabels.Font.Name = "Calibri"
    ch.Axes(xlValue, xlPrimary).TickLabels.Font.Size = 9
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "0 ""чел."""
    ch.Axes(xlValue, xlPrimary).MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
    ch.Axes(xlValue, xlPrimary).MajorGridlines.Format.Line.Weight = 0.75

    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(170, 170, 170)
    ch.Axes(xlValue, xlPrimary).Format.Line.ForeColor.RGB = RGB(170, 170, 170)

    ch.HasDataTable = False

    On Error GoTo 0
End Sub

' ============================================================
' ВСПОМОГАТЕЛЬНОЕ ДЛЯ EXCEL
' ============================================================

Private Function GetUniqueSheetName(ByVal wb As Object, ByVal baseName As String) As String
    Dim cleanName As String
    Dim tryName As String
    Dim n As Long

    cleanName = MakeSafeSheetName(baseName)
    tryName = cleanName

    n = 1

    Do While WorksheetExists(wb, tryName)
        n = n + 1
        tryName = Left$(cleanName, 31 - Len("_" & CStr(n))) & "_" & CStr(n)
    Loop

    GetUniqueSheetName = tryName
End Function

Private Function MakeSafeSheetName(ByVal s As String) As String
    Dim badChars As Variant
    Dim i As Long

    badChars = Array("\", "/", "?", "*", "[", "]", ":")

    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, CStr(badChars(i)), "_")
    Next i

    s = Trim$(s)

    If Len(s) = 0 Then s = "Лист"
    If Len(s) > 31 Then s = Left$(s, 31)

    MakeSafeSheetName = s
End Function

Private Function WorksheetExists(ByVal wb As Object, ByVal sheetName As String) As Boolean
    On Error GoTo NotExists

    Dim sh As Object
    Set sh = wb.Worksheets(sheetName)

    WorksheetExists = True
    Exit Function

NotExists:
    WorksheetExists = False
End Function

Private Function CountMonthsInclusive(ByVal SD0 As Date, ByVal ED0 As Date) As Long
    Dim m As Date
    Dim lastMonth As Date
    Dim cnt As Long

    m = DateSerial(Year(SD0), Month(SD0), 1)
    lastMonth = DateSerial(Year(ED0), Month(ED0), 1)

    Do While m <= lastMonth
        cnt = cnt + 1
        m = DateAdd("m", 1, m)
    Loop

    CountMonthsInclusive = cnt
End Function

Private Function MaxDate(ByVal d1 As Date, ByVal d2 As Date) As Date
    If d1 >= d2 Then
        MaxDate = d1
    Else
        MaxDate = d2
    End If
End Function

Private Function MinDate(ByVal d1 As Date, ByVal d2 As Date) As Date
    If d1 <= d2 Then
        MinDate = d1
    Else
        MinDate = d2
    End If
End Function

Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a <= b Then
        MinLong = a
    Else
        MinLong = b
    End If
End Function

' ============================================================
' РАСЧЁТ РЕСУРСОВ ЗА ОДИН ПРОХОД ПО НАЗНАЧЕНИЯМ
' ============================================================

Private Function CalculatePeriodResourceValues(ByVal tsk As Task, _
                                               ByVal SD0 As Date, _
                                               ByVal adjustedED As Date) As TPeriodResourceValues
    On Error GoTo SafeExit

    Dim result As TPeriodResourceValues
    Dim ASS As Assignment
    Dim res As Resource
    Dim groupKey As String

    For Each ASS In tsk.Assignments
        Set res = ASS.Resource

        If Not res Is Nothing Then
            groupKey = NormalizeGroupKey(res.Group)

            Select Case groupKey
                Case "ЧЕЛОВЕК"
                    result.HumanPlan = result.HumanPlan + _
                        SumAssignmentTimeScale(ASS, SD0, adjustedED, pjAssignmentTimescaledWork, pjTimescaleDays, True)

                    result.HumanFact = result.HumanFact + _
                        SumAssignmentTimeScale(ASS, SD0, adjustedED, pjAssignmentTimescaledActualWork, pjTimescaleDays, True)

                Case "МИМ"
                    result.MimPlan = result.MimPlan + _
                        SumAssignmentTimeScale(ASS, SD0, adjustedED, pjAssignmentTimescaledWork, pjTimescaleDays, True)

                    result.MimFact = result.MimFact + _
                        SumAssignmentTimeScale(ASS, SD0, adjustedED, pjAssignmentTimescaledActualWork, pjTimescaleDays, True)

                Case "ОБЪЕМ", "ОБЬЕМ"
                    result.VolumePlan = result.VolumePlan + _
                        SumAssignmentTimeScale(ASS, SD0, adjustedED, pjAssignmentTimescaledWork, pjTimescaleDays, False)

                    result.VolumeFact = result.VolumeFact + _
                        SumAssignmentTimeScale(ASS, SD0, adjustedED, pjAssignmentTimescaledActualWork, pjTimescaleDays, False)

                    result.VolumeFactCum = result.VolumeFactCum + _
                        SumAssignmentTimeScale(ASS, Int(ActiveProject.ProjectStart), adjustedED, pjAssignmentTimescaledActualWork, pjTimescaleDays, False)

                    result.VolumeFactPast = result.VolumeFactPast + _
                        SumAssignmentTimeScale(ASS, Int(ActiveProject.ProjectStart), SD0, pjAssignmentTimescaledActualWork, pjTimescaleDays, False)
            End Select
        End If
    Next ASS

    CalculatePeriodResourceValues = result
    Exit Function

SafeExit:
    CalculatePeriodResourceValues = result
End Function

Private Function SumAssignmentTimeScale(ByVal ASS As Assignment, _
                                        ByVal startDate As Date, _
                                        ByVal endExclusive As Date, _
                                        ByVal dataType As Long, _
                                        ByVal scaleUnit As Long, _
                                        ByVal convertMinutesToHours As Boolean) As Double
    On Error GoTo SafeExit

    Dim tsData As TimeScaleValues
    Dim tsValue As TimeScaleValue
    Dim total As Double
    Dim v As Double

    If endExclusive <= startDate Then
        SumAssignmentTimeScale = 0
        Exit Function
    End If

    Set tsData = ASS.TimeScaleData(startDate, endExclusive, dataType, scaleUnit)

    If Not tsData Is Nothing Then
        For Each tsValue In tsData
            If tsValue.StartDate >= startDate And tsValue.StartDate < endExclusive Then
                v = ToD(tsValue.Value)
                If v > 0 Then total = total + v
            End If
        Next tsValue
    End If

    If convertMinutesToHours Then
        total = total / 60#
    End If

    SumAssignmentTimeScale = total
    Exit Function

SafeExit:
    SumAssignmentTimeScale = 0
End Function

Private Function NormalizeGroupKey(ByVal s As String) As String
    s = Trim$(CStr(s))
    s = Replace$(s, "ё", "е", , , vbTextCompare)
    s = Replace$(s, "Ё", "Е", , , vbTextCompare)
    NormalizeGroupKey = UCase$(s)
End Function

' ============================================================
' ОСНОВНОЙ РАСЧЁТ ПЕРИОДА В MS PROJECT
' ============================================================

Private Sub RunCalcAndFilter(ByVal SD As Date, ByVal ED As Date, _
                             ByVal applyFilter As Boolean, _
                             ByVal showMsg As Boolean)
    On Error GoTo EH

    If applyFilter Then ResetViewFilter

    Dim SD0 As Date
    Dim ED0 As Date
    Dim adjustedED As Date

    SD0 = Int(SD)
    ED0 = Int(ED)
    adjustedED = ED0 + 1

    Dim tsk As Task
    Dim ASS As Assignment

    Dim s As Date
    Dim F As Date
    Dim dlitt As Double

    Dim p As Double
    Dim stanstav As Double

    Dim allTasks As Long
    Dim periodTasks As Long
    Dim taskCounter As Long
    Dim progressPercentage As Double

    Dim totalTaskDurationDays As Double
    Dim periodDurationDays As Double

    Dim rv As TPeriodResourceValues

    allTasks = 0
    periodTasks = 0

    For Each tsk In ActiveProject.Tasks
        If Not tsk Is Nothing Then
            allTasks = allTasks + 1
            If TaskIntersectsPeriod(tsk, SD0, ED0) Then periodTasks = periodTasks + 1
        End If
    Next tsk

    If allTasks > 0 Then
        taskCounter = 0

        For Each tsk In ActiveProject.Tasks
            If Not tsk Is Nothing Then
                taskCounter = taskCounter + 1

                If taskCounter = 1 Or taskCounter Mod 100 = 0 Or taskCounter = allTasks Then
                    progressPercentage = (taskCounter / allTasks) * 100

                    Application.StatusBar = "Период " & _
                        Format(SD0, "dd.mm.yyyy") & " — " & Format(ED0, "dd.mm.yyyy") & _
                        " | Задача " & taskCounter & " из " & allTasks & _
                        " (" & Round(progressPercentage, 2) & "%)"
                End If

                If tsk.Summary = False Then
                    tsk.Number2 = 0
                    tsk.Number3 = 0
                    tsk.Number6 = 0
                    tsk.Number8 = 0
                    tsk.Number10 = 0
                    tsk.Number11 = 0
                    tsk.Number12 = 0
                    tsk.Number13 = 0
                    tsk.Number17 = 0
                    tsk.Number18 = 0
                    tsk.Cost2 = 0
                    tsk.Cost3 = 0

                    If TaskIntersectsPeriod(tsk, SD0, ED0) Then
                        s = IIf(tsk.Start < SD0, SD0, tsk.Start)
                        F = IIf(tsk.Finish > ED0, ED0, tsk.Finish)

                        totalTaskDurationDays = (Int(tsk.Finish) - Int(tsk.Start) + 1)
                        periodDurationDays = (Int(F) - Int(s) + 1)

                        dlitt = periodDurationDays * 8#

                        If totalTaskDurationDays > 0 Then
                            If tsk.Number5 > 0 Then
                                tsk.Number17 = tsk.Number5 / (totalTaskDurationDays * 8#) * dlitt
                            End If

                            If tsk.Number9 > 0 Then
                                tsk.Number18 = tsk.Number9 / (totalTaskDurationDays * 8#) * dlitt
                            End If
                        End If

                        rv = CalculatePeriodResourceValues(tsk, SD0, adjustedED)

                        tsk.Number6 = Round(rv.HumanPlan, 2)
                        tsk.Number8 = Round(rv.HumanFact, 2)
                        tsk.Number10 = Round(rv.MimPlan, 2)
                        tsk.Number12 = Round(rv.MimFact, 2)

                        If rv.VolumePlan > 0 Then tsk.Number2 = Round(rv.VolumePlan, 2)
                        If rv.VolumeFact > 0 Then tsk.Number3 = Round(rv.VolumeFact, 2)
                        If rv.VolumeFactCum > 0 Then tsk.Number11 = Round(rv.VolumeFactCum, 2)
                        If rv.VolumeFactPast > 0 Then tsk.Number13 = Round(rv.VolumeFactPast, 2)

                        For Each ASS In tsk.Assignments
                            p = GetAssignmentUnitsFactor(ASS)
                            stanstav = GetHourlyRateSimple(ASS.Resource)

                            If Not ASS.Resource Is Nothing Then
                                Select Case NormalizeGroupKey(ASS.Resource.Group)
                                    Case "ЧЕЛОВЕК"
                                        tsk.Cost2 = tsk.Cost2 + stanstav * dlitt * p

                                    Case "МИМ"
                                        tsk.Cost3 = tsk.Cost3 + stanstav * dlitt * p
                                End Select
                            End If
                        Next ASS
                    End If
                End If
            End If
        Next tsk
    End If

    Application.StatusBar = False

    ForceProjectRecalc

    If applyFilter Then ApplyDateRangeFilter SD0, ED0

    If showMsg Then
        MsgBox "Обработка завершена!" & vbCrLf & _
               "Период: " & Format(SD0, "dd.mm.yyyy") & " — " & Format(ED0, "dd.mm.yyyy") & vbCrLf & _
               "Задач в периоде: " & periodTasks, vbInformation
    End If

    Exit Sub

EH:
    Application.StatusBar = False
    MsgBox "Ошибка: " & Err.Number & " — " & Err.Description, vbExclamation
End Sub

Private Function TaskIntersectsPeriod(ByVal t As Task, ByVal SD0 As Date, ByVal ED0 As Date) As Boolean
    On Error GoTo SafeExit

    TaskIntersectsPeriod = (t.Start < ED0 + 1 And t.Finish >= SD0)
    Exit Function

SafeExit:
    TaskIntersectsPeriod = False
End Function

Private Function GetAssignmentUnitsFactor(ByVal ASS As Assignment) As Double
    On Error GoTo SafeExit

    Dim unitsStr As String
    Dim p As Double

    If IsNumeric(ASS.Units) Then
        p = CDbl(ASS.Units)
    Else
        unitsStr = CStr(ASS.Units)
        unitsStr = Replace$(unitsStr, "%", "")
        unitsStr = Replace$(unitsStr, ",", ".")

        If IsNumeric(unitsStr) Then
            p = CDbl(unitsStr) / 100#
        Else
            p = 1#
        End If
    End If

    If p <= 0 Then p = 1#

    GetAssignmentUnitsFactor = p
    Exit Function

SafeExit:
    GetAssignmentUnitsFactor = 1#
End Function

' ============================================================
' ФИЛЬТРАЦИЯ И ФЛАГИ
' ============================================================

Private Sub ResetViewFilter()
    On Error Resume Next

    Application.FilterClear

    If Err.Number <> 0 Then
        Err.Clear
        Application.FilterApply "Все задачи"

        If Err.Number <> 0 Then
            Err.Clear
            Application.FilterApply "All Tasks"
        End If
    End If

    On Error GoTo 0
End Sub

Private Sub ForceProjectRecalc()
    On Error Resume Next

    Application.CalculateProject

    If Err.Number <> 0 Then
        Err.Clear
        Application.CommandBars.ExecuteMso "CalculateProject"
    End If

    DoEvents

    On Error GoTo 0
End Sub

Private Sub ApplyDateRangeFilter(ByVal SD As Date, ByVal ED As Date)
    On Error GoTo FailSoft

    Dim SD0 As Date
    Dim ED0 As Date
    Dim t As Task

    SD0 = Int(SD)
    ED0 = Int(ED)

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            t.Flag20 = False
            t.Flag19 = False
        End If
    Next t

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If Not t.Summary Then
                If TaskIntersectsPeriod(t, SD0, ED0) Then
                    t.Flag20 = True
                End If
            End If
        End If
    Next t

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If Not t.Summary Then
                Dim parent As Task
                Set parent = t.OutlineParent

                If Not parent Is Nothing Then
                    If ParentHasDirectChildWithFlag20(parent) Then
                        t.Flag19 = True
                    Else
                        t.Flag19 = False
                    End If
                Else
                    t.Flag19 = False
                End If
            End If
        End If
    Next t

    On Error Resume Next

    Application.FilterClear
    If Err.Number <> 0 Then Err.Clear

    Application.FilterApply "Флаг20"

    If Err.Number <> 0 Then
        Err.Clear
        Application.FilterApply "Flag20"
    End If

    On Error GoTo 0
    Exit Sub

FailSoft:
    Application.StatusBar = "Не удалось применить фильтр 'Флаг20'. Расчёты выполнены."
End Sub

Private Function ParentHasDirectChildWithFlag20(ByVal parent As Task) As Boolean
    If parent Is Nothing Then
        ParentHasDirectChildWithFlag20 = False
        Exit Function
    End If

    Dim child As Task

    For Each child In parent.OutlineChildren
        If Not child Is Nothing Then
            If child.Flag20 Then
                ParentHasDirectChildWithFlag20 = True
                Exit Function
            End If
        End If
    Next child

    ParentHasDirectChildWithFlag20 = False
End Function

Private Function GetHourlyRateSimple(ByVal r As Resource) As Double
    On Error GoTo SafeExit

    If r Is Nothing Then
        GetHourlyRateSimple = 0
        Exit Function
    End If

    Dim s As String

    s = CStr(r.StandardRate)
    s = Replace$(s, ",", ".")
    s = Replace$(s, "руб", "", , , vbTextCompare)
    s = Replace$(s, "р", "", , , vbTextCompare)
    s = Replace$(s, ChrW(&H20BD), "")
    s = Replace$(s, "RUB", "", , , vbTextCompare)
    s = Replace$(s, "$", "")
    s = Replace$(s, "€", "")
    s = Trim$(s)

    GetHourlyRateSimple = Val(s)
    Exit Function

SafeExit:
    GetHourlyRateSimple = 0
End Function
