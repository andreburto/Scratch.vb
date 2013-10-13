VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Week Days"
   ClientHeight    =   5715
   ClientLeft      =   240
   ClientTop       =   6105
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMonthDays 
      Height          =   840
      ItemData        =   "scratch.frx":0000
      Left            =   120
      List            =   "scratch.frx":0002
      TabIndex        =   24
      Top             =   4800
      Width           =   4455
   End
   Begin VB.CommandButton btnGetFirst 
      Caption         =   "Get First Day Of Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.ComboBox mnuSetMonth 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Choose month"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtSetYear 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton btnTomorrow 
      Caption         =   ">>"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton btnYesterday 
      Appearance      =   0  'Flat
      Caption         =   "<<"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Days (and dates) in the Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4440
      Width           =   4455
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   120
      X2              =   4560
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Start:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   120
      X2              =   4560
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TodayIs = "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblTodayIs 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblShowDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblDaysInMonth 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblDaysInYear 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Days in Month:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Days in Year:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   120
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblYear 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblWeekDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Weekday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   120
      X2              =   4560
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Weekly(1 To 7) As New WeekDays
Private DayNames(1 To 7) As String
Private MonthDays(1 To 12) As Integer
Private MonthNames(1 To 12) As String
Private Monthly(1 To 12) As New Months
Private ThisDay As New WeekDays
Private ThisDate As New TodayIs
Private TodayIs As Double
Private ThisYear As Integer

Private Sub btnGetFirst_Click()
    ' GotoIs is a number that represents the first day of the month in the year
    ' that you want to find.
    Dim GotoIs As Double
    
    ' Dumps the month number
    Dim getMonth As Integer
    getMonth = GetMonthNumber(mnuSetMonth.Text)
    
    ' searchDirection is either "Back" or "Forward"
    Dim searchDirection As String
    
    ' In case the month or year is blank
    If txtSetYear.Text = "" Or getMonth = 0 Then
        MsgBox "Please choose a month and year!", vbOKOnly, "Hey, dingbat..."
        Exit Sub
    End If
    
    ' This will set GotoIs
    GotoIs = SetTodayIs(CInt(txtSetYear.Text), getMonth, 1)
    
    If GotoIs < TodayIs Then
        searchDirection = "Back"
    ElseIf GotoIs > TodayIs Then
        searchDirection = "Forward"
    Else
        MsgBox "You are already there!", vbOKOnly, "Hey, dingbat..."
        Exit Sub
    End If
    
    Do Until GotoIs = TodayIs
        SetToday (searchDirection)
    Loop
    
    If lstMonthDays.ListCount > 0 Then
        lstMonthDays.Clear
    End If
    
    Dim monthDayCount As Integer, DayOfTheWeek As Integer
    DayOfTheWeek = ThisDate.ThisDay
    For monthDayCount = 1 To MonthDays(getMonth)
        lstMonthDays.AddItem (DayNames(DayOfTheWeek) & " (" & monthDayCount & ")")
        DayOfTheWeek = Weekly(DayOfTheWeek).Tomorrow
    Next
End Sub

Private Sub Form_Load()

    ' SET MONTH INFORMATION
    ' January
    MonthNames(1) = "January"
    Monthly(1).LastMonth = 12
    Monthly(1).ThisMonth = 1
    Monthly(1).NextMonth = 2
    
    ' February
    MonthNames(2) = "February"
    Monthly(2).LastMonth = 1
    Monthly(2).ThisMonth = 2
    Monthly(2).NextMonth = 3
    
    ' March
    MonthNames(3) = "March"
    Monthly(3).LastMonth = 2
    Monthly(3).ThisMonth = 3
    Monthly(3).NextMonth = 4
    
    ' April
    MonthNames(4) = "April"
    Monthly(4).LastMonth = 3
    Monthly(4).ThisMonth = 4
    Monthly(4).NextMonth = 5
    
    ' May
    MonthNames(5) = "May"
    Monthly(5).LastMonth = 4
    Monthly(5).ThisMonth = 5
    Monthly(5).NextMonth = 6
    
    ' June
    MonthNames(6) = "June"
    Monthly(6).LastMonth = 5
    Monthly(6).ThisMonth = 6
    Monthly(6).NextMonth = 7
    
    ' July
    MonthNames(7) = "July"
    Monthly(7).LastMonth = 6
    Monthly(7).ThisMonth = 7
    Monthly(7).NextMonth = 8
    
    ' August
    MonthNames(8) = "August"
    Monthly(8).LastMonth = 7
    Monthly(8).ThisMonth = 8
    Monthly(8).NextMonth = 9
    
    ' September
    MonthNames(9) = "September"
    Monthly(9).LastMonth = 8
    Monthly(9).ThisMonth = 9
    Monthly(9).NextMonth = 10
    
    ' October
    MonthNames(10) = "October"
    Monthly(10).LastMonth = 9
    Monthly(10).ThisMonth = 10
    Monthly(10).NextMonth = 11
    
    ' November
    MonthNames(11) = "November"
    Monthly(11).LastMonth = 10
    Monthly(11).ThisMonth = 11
    Monthly(11).NextMonth = 12
    
    ' December
    MonthNames(12) = "December"
    Monthly(12).LastMonth = 11
    Monthly(12).ThisMonth = 12
    Monthly(12).NextMonth = 1

    ' SET DAY INFOMATION
    ' Sunday
    DayNames(1) = "Sunday"
    Weekly(1).Today = 1
    Weekly(1).Yesterday = 7
    Weekly(1).Tomorrow = 2

    ' Monday
    DayNames(2) = "Monday"
    Weekly(2).Today = 2
    Weekly(2).Yesterday = 1
    Weekly(2).Tomorrow = 3

    ' Tuesday
    DayNames(3) = "Tuesday"
    Weekly(3).Today = 3
    Weekly(3).Yesterday = 2
    Weekly(3).Tomorrow = 4

    ' Wednesday
    DayNames(4) = "Wednesday"
    Weekly(4).Today = 4
    Weekly(4).Yesterday = 3
    Weekly(4).Tomorrow = 5

    ' Thursday
    DayNames(5) = "Thursday"
    Weekly(5).Today = 5
    Weekly(5).Yesterday = 4
    Weekly(5).Tomorrow = 6

    ' Friday
    DayNames(6) = "Friday"
    Weekly(6).Today = 6
    Weekly(6).Yesterday = 5
    Weekly(6).Tomorrow = 7

    ' Saturday
    DayNames(7) = "Saturday"
    Weekly(7).Today = 7
    Weekly(7).Yesterday = 6
    Weekly(7).Tomorrow = 1

    ' Sets ThisDate
    ThisDate.ThisDay = WeekDay(Now)
    ThisDate.ThisMonth = Month(Now)
    ThisDate.ThisDate = Day(Now)
    ThisDate.ThisYear = Year(Now)
    
    ' Sets the drop down menu
    Dim monthNum As Integer
    For monthNum = 1 To 12
        mnuSetMonth.AddItem (MonthNames(monthNum))
    Next
    
    ' Sets the middle form
    lblWeekDay = DayNames(ThisDate.ThisDay)
    lblMonth = MonthNames(ThisDate.ThisMonth)
    lblDate = ThisDate.ThisDate
    lblYear = ThisDate.ThisYear
    
    ' Sets starting date form
    lblStart.Caption = MonthNames(ThisDate.ThisMonth) & " " & CStr(ThisDate.ThisDate) _
                       & ", " & CStr(ThisDate.ThisYear)

    ' Sets the bottom forms
    lblDaysInYear.Caption = CStr(CountDays(ThisDate.ThisYear))
    lblDaysInMonth.Caption = CStr(MonthDays(ThisDate.ThisMonth))
    
    ' Sets TodayIs
    TodayIs = SetTodayIs(ThisDate.ThisYear, ThisDate.ThisMonth, ThisDate.ThisDate)
    
    ' Sets ThisDay
    SetToday (ThisDate.ThisDay)
End Sub

Private Sub SetToday(ByVal Direction As String)

    ' If Direction is "Back" go back
    If Direction = "Back" Then

        ' Sets the current week day
        ThisDate.ThisDay = ThisDay.Yesterday
        
        ' Sets the current date
        If CInt(ThisDate.ThisDate - 1) = 0 Then
            ' Sets the current month
            ThisDate.ThisMonth = Monthly(ThisDate.ThisMonth).LastMonth
    
            ' Sets the bottom month form
            lblDaysInMonth.Caption = CStr(MonthDays(ThisDate.ThisMonth))
            
            If ThisDate.ThisMonth = 12 Then
                ' Sets the current year
                ThisDate.ThisYear = ThisDate.ThisYear - 1
                
                ' Sets the bottom year form
                lblDaysInYear.Caption = CStr(CountDays(ThisDate.ThisYear))
            End If
            
            ThisDate.ThisDate = MonthDays(ThisDate.ThisMonth)
        Else
            ThisDate.ThisDate = ThisDate.ThisDate - 1
        End If
            
    ' If Direction is "Forward" go forward
    ElseIf Direction = "Forward" Then

        ' Sets the current week day
        ThisDate.ThisDay = ThisDay.Tomorrow

        ' Sets the current date
        If CInt(ThisDate.ThisDate + 1) > MonthDays(ThisDate.ThisMonth) Then
            ' Sets the current month
            ThisDate.ThisMonth = Monthly(ThisDate.ThisMonth).NextMonth
            
            ' Sets the bottom month form
            lblDaysInMonth.Caption = CStr(MonthDays(ThisDate.ThisMonth))
            
            
            If ThisDate.ThisMonth = 1 Then
                ' Sets the current year
                ThisDate.ThisYear = ThisDate.ThisYear + 1
                
                ' Sets the bottom year form
                lblDaysInYear.Caption = CStr(CountDays(ThisDate.ThisYear))
            End If
            
            ThisDate.ThisDate = 1
        Else
            ThisDate.ThisDate = ThisDate.ThisDate + 1
        End If
    
    End If
        
    ' Sets the information in the middle form
    lblWeekDay = DayNames(ThisDate.ThisDay)
    lblMonth = MonthNames(ThisDate.ThisMonth)
    lblDate = CStr(ThisDate.ThisDate)
    lblYear = CStr(ThisDate.ThisYear)
    
    ' Puts the day in the top field
    lblShowDay.Caption = DayNames(ThisDate.ThisDay)
    
    ' Shifts the week
    ThisDay.Today = Weekly(ThisDate.ThisDay).Today
    ThisDay.Tomorrow = Weekly(ThisDate.ThisDay).Tomorrow
    ThisDay.Yesterday = Weekly(ThisDate.ThisDay).Yesterday
    
    ' Need this as a place holder
    TodayIs = SetTodayIs(ThisDate.ThisYear, ThisDate.ThisMonth, ThisDate.ThisDate)
    lblTodayIs.Caption = CStr(TodayIs)
End Sub

Public Function SetTodayIs(ByVal YR As Integer, ByVal MN As Integer, ByVal DY As Integer) As Double
    Dim tmpToday As String, tmpMN As String, tmpDY As String

    ' Month to string
    If MN < 10 Then
        tmpMN = "0" & CStr(MN)
    Else
        tmpMN = CStr(MN)
    End If
    
    ' Day to string
    If DY < 10 Then
        tmpDY = "0" & CStr(DY)
    Else
        tmpDY = CStr(DY)
    End If

    tmpToday = CStr(YR) & tmpMN & tmpDY

    SetTodayIs = CDbl(tmpToday)
End Function

Public Function GetMonthNumber(ByVal ChosenMonth As String) As Integer
    Dim countMonths As Integer, pickedMonth As Integer
    pickedMonth = 0
    
    For countMonths = 1 To 12
        If MonthNames(countMonths) = ChosenMonth Then
            pickedMonth = countMonths
        End If
    Next
    
    GetMonthNumber = pickedMonth
End Function

Public Function CountDays(ByVal StatedYear As Integer) As Integer
    SetMonthDays (StatedYear)
    
    ' Sets the days in the year
    Dim tmpDays As Integer
    
    For cntMon = 1 To 12
        tmpDays = tmpDays + MonthDays(cntMon)
    Next cntMon

    CountDays = tmpDays
End Function

' Sets days for the months
Public Sub SetMonthDays(ByVal StatedYear As Integer)
    MonthDays(1) = 31
    If IsLeap(StatedYear) = True Then
        MonthDays(2) = 29
    Else
        MonthDays(2) = 28
    End If
    MonthDays(3) = 31
    MonthDays(4) = 30
    MonthDays(5) = 31
    MonthDays(6) = 30
    MonthDays(7) = 31
    MonthDays(8) = 31
    MonthDays(9) = 30
    MonthDays(10) = 31
    MonthDays(11) = 30
    MonthDays(12) = 31
End Sub

Private Function IsLeap(y) As Boolean
    Dim rv As Boolean
    If (y Mod 400) = 0 Then
        rv = True
    ElseIf (y Mod 100) = 0 Then
        rv = False
    ElseIf (y Mod 4) = 0 Then
        rv = True
    Else
        rv = False
    End If
    IsLeap = rv
End Function

Private Sub btnTomorrow_Click()
    SetToday ("Forward")
End Sub

Private Sub btnYesterday_Click()
    SetToday ("Back")
End Sub

Private Sub mnuSetMonth_GotFocus()
    mnuSetMonth.BackColor = &HFFFFC0
End Sub

Private Sub mnuSetMonth_LostFocus()
    mnuSetMonth.BackColor = &HFFFFFF
End Sub

Private Sub txtSetYear_GotFocus()
    txtSetYear.BackColor = &HFFFFC0
End Sub

Private Sub txtSetYear_LostFocus()
    txtSetYear.BackColor = &HFFFFFF
End Sub
