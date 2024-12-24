VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form13 
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
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   16155
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Control Panel"
      Height          =   1335
      Left            =   10320
      TabIndex        =   33
      Top             =   4440
      Width           =   5655
      Begin VB.CommandButton cmdCREATE 
         Caption         =   "&Create"
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdDELETE 
         Caption         =   "&Delete"
         Height          =   615
         Left            =   2400
         TabIndex        =   37
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3600
         TabIndex        =   36
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdUPDATE 
         Caption         =   "&Update"
         Height          =   615
         Left            =   1200
         TabIndex        =   35
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdEXIT 
         Caption         =   "E&xit"
         Height          =   615
         Left            =   4560
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Navigation Panel"
      Height          =   1335
      Left            =   0
      TabIndex        =   28
      Top             =   4440
      Width           =   5655
      Begin VB.CommandButton cmdNEXT 
         Caption         =   "&Next"
         Height          =   615
         Left            =   3195
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdLAST 
         Caption         =   "&Last"
         Height          =   615
         Left            =   4320
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdFIRST 
         Caption         =   "&First"
         Height          =   615
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdPREV 
         Caption         =   "&Previous"
         Height          =   615
         Left            =   1485
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DEPARTMENT ENTRY FORM"
      Height          =   3615
      Left            =   0
      TabIndex        =   15
      Top             =   720
      Width           =   5655
      Begin VB.TextBox txtBUDGET 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """?"" #,##0.00;(""?"" #,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   20
         Top             =   2850
         Width           =   2415
      End
      Begin VB.TextBox txtMANAGER 
         Height          =   495
         Left            =   3000
         TabIndex        =   19
         Top             =   2235
         Width           =   2415
      End
      Begin VB.TextBox txtLOCATION 
         Height          =   495
         Left            =   3000
         TabIndex        =   18
         Top             =   1605
         Width           =   2415
      End
      Begin VB.TextBox txtDEPT_NM 
         Height          =   495
         Left            =   3000
         TabIndex        =   17
         Top             =   990
         Width           =   2415
      End
      Begin VB.TextBox txtDEPT_ID 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Budget"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   2880
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Manager"
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Department Name"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Width           =   2220
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
         TabIndex        =   22
         Top             =   360
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
         Left            =   2640
         TabIndex        =   21
         Top             =   960
         Width           =   120
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "SEARCH BY"
      Height          =   3615
      Left            =   5760
      TabIndex        =   4
      Top             =   720
      Width           =   10095
      Begin VB.ComboBox cmbSearchMANAGER 
         Height          =   495
         Left            =   6240
         TabIndex        =   14
         Text            =   " Manager"
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox cmbSearchLOCATION 
         Height          =   495
         Left            =   4560
         TabIndex        =   13
         Text            =   " Location"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbSearchDEPT_NM 
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Text            =   " Department Name"
         Top             =   480
         Width           =   2775
      End
      Begin VB.ListBox listSearchDEPT_NM 
         Height          =   2085
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   960
         Width           =   2775
      End
      Begin VB.ListBox listSearchLOCATION 
         Height          =   2085
         Left            =   4560
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.ListBox listSearchMANAGER 
         Height          =   2085
         Left            =   6240
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cmbSearchDEPT_ID 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Text            =   "ID"
         Top             =   480
         Width           =   1455
      End
      Begin VB.ListBox listSearchDEPT_ID 
         Height          =   2085
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox listSearchBUDGET 
         Height          =   2085
         Left            =   8040
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cmbSearchBUDGET 
         Height          =   495
         Left            =   8040
         TabIndex        =   5
         Text            =   " Budget"
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Display Panel"
      Height          =   1335
      Left            =   5760
      TabIndex        =   1
      Top             =   4440
      Width           =   4455
      Begin VB.CommandButton cmdDeptSearch 
         Caption         =   "Search P&anel"
         Height          =   615
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   1845
      End
      Begin VB.CommandButton cmdDispDept 
         Caption         =   "Display &Records"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2200
      End
   End
   Begin MSDataGridLib.DataGrid DataGridDept 
      Bindings        =   "Form13.frx":94CA
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   4895
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   27
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
         Size            =   12
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3480
      Top             =   6720
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=PRJ2432N/PRJ2432N;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=PRJ2432N/PRJ2432N;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM DEPARTMENT"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   6120
      TabIndex        =   39
      Top             =   0
      Width           =   3210
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      X1              =   3720
      X2              =   5880
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   4
      X1              =   9480
      X2              =   11520
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   8040
      X2              =   8040
      Y1              =   0
      Y2              =   8640
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text4_Change()

End Sub
