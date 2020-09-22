VERSION 5.00
Object = "*\A..\Calendar.vbp"
Begin VB.Form frmDemoS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar Style Demo"
   ClientHeight    =   7332
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10344
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7332
   ScaleWidth      =   10344
   StartUpPosition =   2  'CenterScreen
   Begin CalendarOcx.Calendar calDemo4 
      Height          =   3732
      Left            =   4800
      TabIndex        =   3
      Top             =   3480
      Width           =   5412
      _ExtentX        =   9546
      _ExtentY        =   6583
      CellDayOfYearForeColor=   16711680
      CellDayOfYearStyle=   5
      CellDaysStyle   =   5
      CellForeColorSunday=   255
      CellForeColorMonday=   12583104
      CellForeColorTuesday=   12583104
      CellForeColorWednesday=   12583104
      CellForeColorThursday=   12583104
      CellForeColorFriday=   12583104
      CellForeColorSaturday=   16711680
      CellHeaderStyle =   5
      CellOtherMonthForeColor=   14737632
      CellOtherMonthStyle=   5
      CellOtherMonthView=   -1  'True
      CellSelectForeColor=   16711680
      CellSelectHeaderForeColor=   65535
      CellSelectStyle =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16711680
      Language        =   5
      ShowInfoBar     =   0
      ShowNavigationBar=   0
      ShowToolTipText =   0   'False
   End
   Begin CalendarOcx.Calendar calDemo3 
      Height          =   3732
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   4452
      _ExtentX        =   7853
      _ExtentY        =   6583
      Appearance      =   1
      CellDaysBackColor=   16777215
      CellDaysStyle   =   6
      CellSelectBackColor=   16777215
      CellSelectHeaderForeColor=   0
      CellSelectStyle =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FrameStyle      =   2
      GridColor       =   16711680
      GridStyle       =   3
      Language        =   4
      Hemisphere      =   1
      LockInfoBar     =   -1  'True
      ShowInfoBar     =   1
      ShowNavigationBar=   0
   End
   Begin CalendarOcx.Calendar calDemo2 
      Height          =   3132
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   5412
      _ExtentX        =   9546
      _ExtentY        =   5525
      Appearance      =   1
      ArrowColor      =   16711680
      BorderStyle     =   1
      CellDaysBackColor=   16777215
      CellDaysStyle   =   6
      CellSelectBackColor=   16777215
      CellSelectHeaderForeColor=   0
      CellSelectStyle =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FrameStyle      =   2
      GridColor       =   16711680
      GridStyle       =   4
      LabelBackColor  =   14737632
      LabelBorderStyle=   0
      LabelForeColor  =   12582912
      Language        =   3
      ShowInfoBar     =   0
      ShowNavigationBar=   1
   End
   Begin CalendarOcx.Calendar calDemo1 
      Height          =   3132
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4452
      _ExtentX        =   7853
      _ExtentY        =   5525
      CellSelectHeaderForeColor=   0
      CellSelectStyle =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16711680
      SelectedDayMark =   0   'False
      SelectionType   =   1
   End
End
Attribute VB_Name = "frmDemoS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

