VERSION 5.00
Object = "*\A..\Calendar.vbp"
Begin VB.Form frmDemoM 
   Caption         =   "Calendar Markings Demo"
   ClientHeight    =   4692
   ClientLeft      =   132
   ClientTop       =   -180
   ClientWidth     =   9972
   LinkTopic       =   "Form1"
   ScaleHeight     =   4692
   ScaleWidth      =   9972
   StartUpPosition =   2  'CenterScreen
   Begin CalendarOcx.Calendar calDemo 
      Height          =   4068
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   4572
      _ExtentX        =   8065
      _ExtentY        =   7176
      ArrowColor      =   12582912
      BackColor       =   14737632
      ButtonColor     =   12632256
      ButtonGradientColor=   16777215
      ButtonGradientStyle=   4
      CellDayOfYearForeColor=   12583104
      CellDayOfYearStyle=   0
      CellForeColorSunday=   255
      CellForeColorMonday=   8388736
      CellForeColorTuesday=   8388736
      CellForeColorWednesday=   8388736
      CellForeColorThursday=   8388736
      CellForeColorFriday=   8388736
      CellForeColorSaturday=   12582912
      CellHeaderStyle =   0
      CellOtherMonthForeColor=   8421504
      CellOtherMonthView=   -1  'True
      CellSelectForeColor=   12582912
      CellSelectHeaderForeColor=   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FrameStyle      =   1
      FrameColor      =   8421504
      GradientColor   =   16777215
      GradientStyle   =   4
      GridColor       =   12632256
      GridStyle       =   2
      LabelBackStyle  =   0
      LabelBorderStyle=   0
      LabelFontBold   =   -1  'True
      LabelForeColor  =   8388736
      ShowInfoBar     =   1
      ShowNavigationBar=   1
      WeekDayViewChar =   2
      WeekNumberForeColor=   12582912
   End
   Begin VB.Frame fraSettings 
      Caption         =   "Some Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   4572
      Left            =   4920
      TabIndex        =   3
      Top             =   0
      Width           =   4932
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Caption         =   "ShowToolTipText:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   4692
      End
      Begin VB.ComboBox cmbGrid 
         ForeColor       =   &H00800080&
         Height          =   288
         ItemData        =   "frmDemoM.frx":0000
         Left            =   120
         List            =   "frmDemoM.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2280
         Width           =   2292
      End
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Caption         =   "ShowNavigationBar:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   4692
      End
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Caption         =   "ShowInfoBar:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   3480
         Width           =   4692
      End
      Begin VB.CheckBox chkOptions 
         Alignment       =   1  'Right Justify
         Caption         =   "ShowOtherMonths:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   3840
         Width           =   4692
      End
      Begin VB.ComboBox cmbGradient 
         ForeColor       =   &H00800080&
         Height          =   288
         ItemData        =   "frmDemoM.frx":004A
         Left            =   120
         List            =   "frmDemoM.frx":005D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Width           =   2292
      End
      Begin VB.ComboBox cmbButtonGradient 
         ForeColor       =   &H00800080&
         Height          =   288
         ItemData        =   "frmDemoM.frx":009D
         Left            =   120
         List            =   "frmDemoM.frx":00B0
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   2292
      End
      Begin VB.ComboBox cmbLanguage 
         ForeColor       =   &H00800080&
         Height          =   288
         ItemData        =   "frmDemoM.frx":00F0
         Left            =   2520
         List            =   "frmDemoM.frx":0106
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   2292
      End
      Begin VB.ComboBox cmbFrame 
         ForeColor       =   &H00800080&
         Height          =   288
         ItemData        =   "frmDemoM.frx":013C
         Left            =   120
         List            =   "frmDemoM.frx":0149
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   2292
      End
      Begin VB.ComboBox cmbWeekDay 
         ForeColor       =   &H00800080&
         Height          =   288
         ItemData        =   "frmDemoM.frx":0160
         Left            =   2520
         List            =   "frmDemoM.frx":0162
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   2292
      End
      Begin VB.ComboBox cmbDateFormat 
         ForeColor       =   &H00800080&
         Height          =   288
         ItemData        =   "frmDemoM.frx":0164
         Left            =   2520
         List            =   "frmDemoM.frx":0171
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   2292
      End
      Begin VB.ComboBox cmbHemisphere 
         ForeColor       =   &H00800080&
         Height          =   288
         ItemData        =   "frmDemoM.frx":0199
         Left            =   2520
         List            =   "frmDemoM.frx":01A3
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2280
         Width           =   2292
      End
      Begin VB.Label lblSetting 
         AutoSize        =   -1  'True
         Caption         =   "GridStyle:"
         ForeColor       =   &H00C00000&
         Height          =   192
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   696
      End
      Begin VB.Label lblSetting 
         AutoSize        =   -1  'True
         Caption         =   "GradientType:"
         ForeColor       =   &H00C00000&
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1032
      End
      Begin VB.Label lblSetting 
         AutoSize        =   -1  'True
         Caption         =   "ButtonGradientType:"
         ForeColor       =   &H00C00000&
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1476
      End
      Begin VB.Label lblSetting 
         AutoSize        =   -1  'True
         Caption         =   "Language:"
         ForeColor       =   &H00C00000&
         Height          =   312
         Index           =   3
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   768
      End
      Begin VB.Label lblSetting 
         AutoSize        =   -1  'True
         Caption         =   "FrameStyle:"
         ForeColor       =   &H00C00000&
         Height          =   192
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   864
      End
      Begin VB.Label lblSetting 
         AutoSize        =   -1  'True
         Caption         =   "FirstWeekDay:"
         ForeColor       =   &H00C00000&
         Height          =   192
         Index           =   5
         Left            =   2520
         TabIndex        =   17
         Top             =   840
         Width           =   1068
      End
      Begin VB.Label lblSetting 
         AutoSize        =   -1  'True
         Caption         =   "DateFormat:"
         ForeColor       =   &H00C00000&
         Height          =   192
         Index           =   6
         Left            =   2520
         TabIndex        =   16
         Top             =   1440
         Width           =   888
      End
      Begin VB.Label lblSetting 
         AutoSize        =   -1  'True
         Caption         =   "Hemisphere:"
         ForeColor       =   &H00C00000&
         Height          =   192
         Index           =   7
         Left            =   2520
         TabIndex        =   15
         Top             =   2040
         Width           =   924
      End
      Begin VB.Image imgLine 
         BorderStyle     =   1  'Fixed Single
         Height          =   12
         Left            =   120
         Top             =   2880
         Width           =   4692
      End
   End
   Begin VB.ComboBox cmbMonth 
      ForeColor       =   &H00800080&
      Height          =   288
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   45
      Width           =   1092
   End
   Begin VB.ComboBox cmbYear 
      ForeColor       =   &H00800080&
      Height          =   288
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Width           =   960
   End
   Begin VB.Image imgBorder 
      BorderStyle     =   1  'Fixed Single
      Height          =   4116
      Left            =   96
      Top             =   456
      Width           =   4620
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "lblDate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   732
   End
   Begin VB.Menu mnuMarkings 
      Caption         =   "Markings"
      Visible         =   0   'False
      Begin VB.Menu mnuMark 
         Caption         =   "Mark"
         Begin VB.Menu mnuMarkType 
            Caption         =   "Type 1"
            Index           =   0
         End
         Begin VB.Menu mnuMarkType 
            Caption         =   "Type 2"
            Index           =   1
         End
         Begin VB.Menu mnuMarkType 
            Caption         =   "Type 3"
            Index           =   2
         End
         Begin VB.Menu mnuMarkType 
            Caption         =   "Type 4"
            Index           =   3
         End
         Begin VB.Menu mnuMarkType 
            Caption         =   "Type 5"
            Index           =   4
         End
      End
      Begin VB.Menu mnuDemark 
         Caption         =   "Demark"
         Begin VB.Menu mnuDemarkType 
            Caption         =   "Type 1"
            Index           =   0
         End
         Begin VB.Menu mnuDemarkType 
            Caption         =   "Type 2"
            Index           =   1
         End
         Begin VB.Menu mnuDemarkType 
            Caption         =   "Type 3"
            Index           =   2
         End
         Begin VB.Menu mnuDemarkType 
            Caption         =   "Type 4"
            Index           =   3
         End
         Begin VB.Menu mnuDemarkType 
            Caption         =   "Type 5"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "frmDemoM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Demo program for Calendar Control
'
'Author Ben Vonk
'20-08-2004 First version
'29-10-2005 Second version (based on Stefaan Casier's 'Owner Drawn Calendar Control' at http://www.codeguru.com/vb/controls/vb_othctrl/ocxcontrols/article.php/c1521/)

' This example demonstrates how Calendar control can be used.
' It shows how to (de)mark only selected calendar days, using a popup menu.
' The marking-changes you make in this demo are not stored.

Option Explicit
                        ' in case you want to detect changes elsewhere
Dim Changes As Boolean  ' not really used in this program

Private Sub DoMarkers(ByVal Index As Integer, ByVal Status As Boolean)

Dim intDay As Integer

   With calDemo
      For intDay = 1 To .GetMonthDays
         ' day, type, on/off
         ' here comes code that changes your data
         If .IsDaySel(intDay) Then Call .DayMarking(intDay, Index, Status)
      Next 'intDay
      
      .Refresh
      Changes = True
   End With

End Sub

Private Sub FillDate()

   With calDemo
      lblDate.Caption = Format(DateSerial(.CalYear, .CalMonth, .CalDay), cmbDateFormat.Text)
   End With

End Sub

Private Sub calDemo_DayClick(Button As Integer, Shift As Integer, IsDay As Integer, IsMonth As Integer, Cancel As Boolean)

   If Button = vbRightButton Then
      PopupMenu mnuMarkings
      Cancel = True  ' cancel normal action: (de-)selection of cell
   End If

End Sub

Private Sub chkOptions_Click(Index As Integer)

   With calDemo
      Select Case Index
         Case 0
            .ShowNavigationBar = IIf(chkOptions.Item(Index).Value, Small, Off)
            
         Case 1
            .ShowInfoBar = IIf(chkOptions.Item(Index).Value, Small, Off)
            
         Case 2
            .CellOtherMonthView = CBool(chkOptions.Item(Index).Value)
            
         Case 3
            .ShowToolTipText = CBool(chkOptions.Item(Index).Value)
      End Select
   End With

End Sub

Private Sub cmbButtonGradient_Click()

   calDemo.ButtonGradientStyle = cmbButtonGradient.ListIndex

End Sub

Private Sub cmbDateFormat_Click()

   calDemo.DateFormat = cmbDateFormat.ListIndex
   
   Call FillDate

End Sub

Private Sub cmbFrame_Click()

   calDemo.FrameStyle = cmbFrame.ListIndex

End Sub

Private Sub cmbGradient_Click()

   calDemo.GradientStyle = cmbGradient.ListIndex

End Sub

Private Sub cmbGrid_Click()

   calDemo.GridStyle = cmbGrid.ListIndex

End Sub

Private Sub cmbHemisphere_Click()

   calDemo.Hemisphere = cmbHemisphere.ListIndex
   
End Sub

Private Sub cmbLanguage_Click()

   calDemo.Language = cmbLanguage.ListIndex

End Sub

Private Sub cmbMonth_Click()

   calDemo.CalMonth = cmbMonth.ListIndex + 1
   
   Call FillDate

End Sub

Private Sub cmbWeekDay_Click()

   calDemo.FirstWeekDay = cmbWeekDay.ListIndex

End Sub

Private Sub cmbYear_Click()

   calDemo.CalYear = cmbYear.List(cmbYear.ListIndex)
   
   Call FillDate

End Sub

Private Sub Form_Load()

Dim intCount As Integer

   frmDemoS.Show
   DoEvents
   Height = 5472
   intCount = Val(Format(Now, "yyyy"))
    
   For intCount = intCount - 6 To intCount + 6
      cmbYear.AddItem Str(intCount)
   Next 'intCount
   
   With calDemo
      .Locked = True
      cmbYear.ListIndex = 6
      
      For intCount = 1 To 12
         cmbMonth.AddItem calDemo.GetMonthName(intCount)
      Next 'intCount
      
      cmbMonth.ListIndex = Month(Now) - 1
      
      For intCount = 1 To 7
         cmbWeekDay.AddItem .GetWeekDayName(intCount)
      Next 'intCount
      
      cmbWeekDay.ListIndex = .FirstWeekDay
      cmbButtonGradient.ListIndex = .ButtonGradientStyle
      cmbGradient.ListIndex = .GradientStyle
      cmbFrame.ListIndex = .FrameStyle
      cmbGrid.ListIndex = .GridStyle
      cmbLanguage.ListIndex = .Language
      cmbDateFormat.ListIndex = .DateFormat
      cmbHemisphere.ListIndex = .Hemisphere
      chkOptions.Item(0).Value = .ShowNavigationBar
      chkOptions.Item(1).Value = .ShowInfoBar
      chkOptions.Item(2).Value = Abs(.CellOtherMonthView)
      chkOptions.Item(3).Value = Abs(.ShowToolTipText)
      .Locked = False
   End With
   
   Call FillDate

End Sub

Private Sub Form_Resize()

   If WindowState = 1 Then Exit Sub
   
   If WindowState = 0 Then
      If Width < 10068 Then Width = 10068
      If Height < 5472 Then Height = 5472
   End If
   
   With calDemo
      fraSettings.Left = ScaleWidth - fraSettings.Width - 120
      .Width = fraSettings.Left - 300
      .Height = ScaleHeight - calDemo.Top - 120
      imgBorder.Left = .Left - Screen.TwipsPerPixelX * 2
      imgBorder.Top = .Top - Screen.TwipsPerPixelY * 2
      imgBorder.Width = .Width + Screen.TwipsPerPixelX * 4
      imgBorder.Height = .Height + Screen.TwipsPerPixelY * 4
   End With

End Sub

Private Sub mnuDemarkType_Click(Index As Integer)

   Call DoMarkers(Index + 1, False)

End Sub

Private Sub mnuMarkType_Click(Index As Integer)

   Call DoMarkers(Index + 1, True)

End Sub

Private Sub calDemo_SelChanged()

   Call FillDate

End Sub

