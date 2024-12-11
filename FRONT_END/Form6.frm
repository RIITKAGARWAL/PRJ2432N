VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDepartment 
   Caption         =   "Department "
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
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   16155
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Display Panel"
      Height          =   1335
      Left            =   5880
      TabIndex        =   35
      Top             =   4440
      Width           =   4455
      Begin VB.CommandButton cmdDispDept 
         Caption         =   "&Display Records"
         Height          =   615
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   2200
      End
      Begin VB.CommandButton cmdDeptSearch 
         Caption         =   "Search &Panel"
         Height          =   615
         Left            =   2400
         TabIndex        =   36
         Top             =   480
         Width           =   1845
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "SEARCH BY"
      Height          =   3615
      Left            =   5880
      TabIndex        =   24
      Top             =   720
      Width           =   10095
      Begin VB.ListBox List1 
         Height          =   2085
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   32
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Text            =   " Department Id"
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox List4 
         Height          =   2085
         Left            =   7560
         Style           =   1  'Checkbox
         TabIndex        =   30
         Top             =   960
         Width           =   2415
      End
      Begin VB.ListBox List3 
         Height          =   2085
         Left            =   5400
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   960
         Width           =   2055
      End
      Begin VB.ListBox List2 
         Height          =   2085
         Left            =   2520
         Style           =   1  'Checkbox
         TabIndex        =   28
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         Height          =   495
         Left            =   2520
         TabIndex        =   27
         Text            =   " Department Name"
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox Combo3 
         Height          =   495
         Left            =   5400
         TabIndex        =   26
         Text            =   " Location"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox Combo4 
         Height          =   495
         Left            =   7560
         TabIndex        =   25
         Text            =   " Manager"
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DEPARTMENT ENTRY FORM"
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   3000
         TabIndex        =   17
         Top             =   990
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   3000
         TabIndex        =   16
         Top             =   1605
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   3000
         TabIndex        =   15
         Top             =   2235
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   3000
         TabIndex        =   14
         Top             =   2850
         Width           =   2415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   525
         Left            =   2640
         TabIndex        =   34
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   525
         Left            =   2280
         TabIndex        =   33
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department Id"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label lblDEPT_ID 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   3000
         TabIndex        =   22
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Department Name"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Manager"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Budget"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   2880
         Width           =   840
      End
   End
   Begin MSDataGridLib.DataGrid DataGridDept 
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   4471
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
   Begin VB.Frame Frame2 
      Caption         =   "Navigation Panel"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   5655
      Begin VB.CommandButton Command10 
         Caption         =   "&Previous"
         Height          =   615
         Left            =   1485
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&First"
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Last"
         Height          =   615
         Left            =   4320
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Next"
         Height          =   615
         Left            =   3195
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control Panel"
      Height          =   1335
      Left            =   10440
      TabIndex        =   0
      Top             =   4440
      Width           =   5655
      Begin VB.CommandButton Command5 
         Caption         =   "&Exit"
         Height          =   615
         Left            =   4560
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Update"
         Height          =   615
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Save"
         Height          =   615
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Delete"
         Height          =   615
         Left            =   2400
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Create"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      X1              =   9600
      X2              =   11640
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      X1              =   3840
      X2              =   6000
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   6240
      TabIndex        =   12
      Top             =   0
      Width           =   3210
   End
End
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toogle As Boolean
Private Sub Form_load()
toogle = True
End Sub
Private Sub cmdDispDept_Click()
toogle = Not toogle
If toogle = False Then
DataGridDept.Visible = True
cmdDispDept.Caption = "&Back"
Else
DataGridDept.Visible = False
cmdDispDept.Caption = "&Display Records"

End If

End Sub
Private Sub cmdDeptSearch_Click()
toogle = False

DataGridDept.Visible = True
cmdDispDept.Caption = "&Back"
End Sub
