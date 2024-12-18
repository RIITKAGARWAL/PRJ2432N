VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPosition 
   Caption         =   "Position"
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
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   16155
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Display Panel"
      Height          =   1335
      Left            =   10560
      TabIndex        =   52
      Top             =   6480
      Width           =   5175
      Begin VB.CommandButton cmdPosSearch 
         Caption         =   "Search &Panel"
         Height          =   615
         Left            =   2520
         TabIndex        =   54
         Top             =   480
         Width           =   2200
      End
      Begin VB.CommandButton cmdDisp 
         Caption         =   "&Display Records"
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   480
         Width           =   2200
      End
   End
   Begin VB.Frame frmConPnl 
      Caption         =   "Control Panel"
      Height          =   1335
      Left            =   7800
      TabIndex        =   23
      Top             =   4440
      Width           =   7335
      Begin VB.CommandButton Command1 
         Caption         =   "&Create"
         Height          =   615
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Delete"
         Height          =   615
         Left            =   3000
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Save"
         Height          =   615
         Left            =   4440
         TabIndex        =   26
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Update"
         Height          =   615
         Left            =   1560
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Exit"
         Height          =   615
         Left            =   5880
         TabIndex        =   24
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame frmNavPnl 
      Caption         =   "Navigation Panel"
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   7335
      Begin VB.CommandButton Command7 
         Caption         =   "&Next"
         Height          =   615
         Left            =   3675
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Last"
         Height          =   615
         Left            =   5400
         TabIndex        =   21
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&First"
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Previous"
         Height          =   615
         Left            =   1965
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame frmPosEntry 
      Caption         =   "POSITION ENTRY FORM"
      Height          =   3615
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   14055
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   11040
         TabIndex        =   50
         Top             =   2355
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   11040
         TabIndex        =   48
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   11040
         TabIndex        =   46
         Top             =   1830
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   11040
         TabIndex        =   44
         Top             =   1290
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   11040
         TabIndex        =   42
         Top             =   765
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   11040
         TabIndex        =   40
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox Combo6 
         Height          =   495
         Left            =   3000
         TabIndex        =   33
         Text            =   " Entry"
         Top             =   2280
         Width           =   2415
      End
      Begin VB.ComboBox Combo5 
         Height          =   495
         Left            =   3000
         TabIndex        =   31
         Text            =   " Department Id"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Professional Development (PDA)"
         Height          =   375
         Left            =   6480
         TabIndex        =   51
         Top             =   2430
         Width           =   4020
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Special Allowance"
         Height          =   375
         Left            =   6480
         TabIndex        =   49
         Top             =   2880
         Width           =   2205
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Medical Allowance (MA)"
         Height          =   375
         Left            =   6480
         TabIndex        =   47
         Top             =   1965
         Width           =   3030
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Conveyance Allowance (TA)"
         Height          =   375
         Left            =   6480
         TabIndex        =   45
         Top             =   1395
         Width           =   3465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Dearness Allowance (DA)"
         Height          =   375
         Left            =   6480
         TabIndex        =   43
         Top             =   810
         Width           =   3105
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "House Rent Allowance (HRA)"
         Height          =   375
         Left            =   6480
         TabIndex        =   41
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label11 
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
         Left            =   1560
         TabIndex        =   39
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label Label10 
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
         TabIndex        =   38
         Top             =   1560
         Width           =   120
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
         Left            =   1800
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   360
         Width           =   120
      End
      Begin VB.Label lblDEPT_ID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3000
         TabIndex        =   32
         Top             =   1095
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Basic Pay"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Position Level"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Width           =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Position Name"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Position Id"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department Id"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1785
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "SEARCH BY"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   10095
      Begin VB.ComboBox Combo7 
         Height          =   495
         Left            =   8280
         TabIndex        =   35
         Text            =   " BasicPay"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List5 
         Height          =   2085
         Left            =   8280
         Style           =   1  'Checkbox
         TabIndex        =   34
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         Height          =   495
         Left            =   6840
         TabIndex        =   8
         Text            =   " Level"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   495
         Left            =   2520
         TabIndex        =   7
         Text            =   " PositionID"
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   495
         Left            =   4560
         TabIndex        =   6
         Text            =   " PositionName"
         Top             =   480
         Width           =   2175
      End
      Begin VB.ListBox List2 
         Height          =   2085
         Left            =   4560
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.ListBox List3 
         Height          =   2085
         Left            =   2520
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.ListBox List4 
         Height          =   2085
         Left            =   6840
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Text            =   " DepartmentID"
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Height          =   2085
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
   End
   Begin MSDataGridLib.DataGrid DataGridPos 
      Height          =   3495
      Left            =   360
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   6165
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "POSITION"
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
      Left            =   6600
      TabIndex        =   30
      Top             =   0
      Width           =   2190
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00008000&
      BorderWidth     =   4
      X1              =   4320
      X2              =   6480
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      BorderWidth     =   4
      X1              =   8880
      X2              =   10920
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   11400
      X2              =   12240
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      X1              =   14040
      X2              =   14880
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12480
      TabIndex        =   29
      Top             =   7920
      Width           =   1395
   End
End
Attribute VB_Name = "frmPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toogle As Boolean


Private Sub cmdPosSearch_Click()
toogle = False
frmPosEntry.Visible = False
frmConPnl.Visible = False
frmNavPnl.Visible = False
DataGridPos.Visible = True
cmdDisp.Caption = "&Back"
End Sub

Private Sub Form_Load()
toogle = True
End Sub
Private Sub cmdDisp_Click()
toogle = Not toogle
If toogle = False Then
frmPosEntry.Visible = False
frmConPnl.Visible = False
frmNavPnl.Visible = False
DataGridPos.Visible = True
cmdDisp.Caption = "&Back"

Else

frmPosEntry.Visible = True
frmConPnl.Visible = True
frmNavPnl.Visible = True
DataGridPos.Visible = False
cmdDisp.Caption = "&Display Records"

End If

End Sub
