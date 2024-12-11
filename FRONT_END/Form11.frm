VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBank 
   Caption         =   "Department & Position"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16155
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   16155
   WindowState     =   2  'Maximized
   Begin VB.Frame frmSearch 
      Caption         =   "SEARCH BY"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   15855
      Begin VB.ListBox List14 
         Height          =   1275
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox Combo14 
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Text            =   " Age"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ListBox List13 
         Height          =   1275
         Left            =   4680
         Style           =   1  'Checkbox
         TabIndex        =   27
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox Combo13 
         Height          =   495
         Left            =   4680
         TabIndex        =   26
         Text            =   " Type"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ListBox List12 
         Height          =   1275
         Left            =   6360
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ComboBox Combo12 
         Height          =   495
         Left            =   6360
         TabIndex        =   24
         Text            =   " Aadhar"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ListBox List11 
         Height          =   1275
         Left            =   8400
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox Combo11 
         Height          =   495
         Left            =   8400
         TabIndex        =   22
         Text            =   " PAN"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ListBox List10 
         Height          =   1275
         Left            =   12360
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox Combo10 
         Height          =   495
         Left            =   12360
         TabIndex        =   20
         Text            =   " Date of Joining"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ListBox List9 
         Height          =   1275
         Left            =   13800
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo9 
         Height          =   495
         Left            =   13800
         TabIndex        =   18
         Text            =   " Status"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List8 
         Height          =   1275
         Left            =   1560
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox Combo8 
         Height          =   495
         Left            =   1560
         TabIndex        =   16
         Text            =   " Gender"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ListBox List7 
         Height          =   1275
         Left            =   3000
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox Combo6 
         Height          =   495
         Left            =   3000
         TabIndex        =   14
         Text            =   "Empty"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ListBox List6 
         Height          =   1275
         Left            =   4560
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox Combo5 
         Height          =   495
         Left            =   4560
         TabIndex        =   12
         Text            =   " EmployeeId"
         Top             =   480
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Height          =   1275
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Text            =   " DepartmentID"
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox List4 
         Height          =   1275
         Left            =   9240
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.ListBox List3 
         Height          =   1275
         Left            =   2520
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
      Begin VB.ListBox List2 
         Height          =   1275
         Left            =   6600
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Height          =   495
         Left            =   6600
         TabIndex        =   5
         Text            =   " EmployeeName"
         Top             =   480
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
         Height          =   495
         Left            =   2520
         TabIndex        =   4
         Text            =   " PositionID"
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox Combo4 
         Height          =   495
         Left            =   9240
         TabIndex        =   3
         Text            =   " Level"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ListBox List5 
         Height          =   1275
         Left            =   10680
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo7 
         Height          =   495
         Left            =   10680
         TabIndex        =   1
         Text            =   " BasicPay"
         Top             =   480
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGridEmp 
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Visible         =   0   'False
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   8040
      X2              =   8040
      Y1              =   0
      Y2              =   8640
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
