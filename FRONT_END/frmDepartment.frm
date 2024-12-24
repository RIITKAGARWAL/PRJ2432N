VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
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
   Icon            =   "frmDepartment.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   16155
   WindowState     =   2  'Maximized
   Begin VB.Frame frmControls 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   12240
      TabIndex        =   15
      Top             =   720
      Width           =   3735
      Begin VB.Frame frameControlPanel 
         Caption         =   "Control Panel"
         Height          =   3855
         Left            =   960
         TabIndex        =   30
         Top             =   2760
         Width           =   1935
         Begin VB.CommandButton cmdEXIT 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            ToolTipText     =   "Closes the Form"
            Top             =   3360
            Width           =   1455
         End
         Begin VB.CommandButton cmdUPDATE 
            Caption         =   "&Update"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   36
            ToolTipText     =   "To Change or Correct Existing Data"
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdSAVE 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   35
            ToolTipText     =   "To Save or Discard the Changes"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdDELETE 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   34
            ToolTipText     =   "To Delete a Record"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCREATE 
            Caption         =   "&Create"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            ToolTipText     =   "To Create a New Record"
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdPrintAll 
            Caption         =   "Print All"
            Height          =   375
            Left            =   240
            TabIndex        =   32
            ToolTipText     =   "Complete Report will be generated"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton cmdDisplay 
            Caption         =   "D&isplay"
            Height          =   375
            Left            =   240
            TabIndex        =   31
            ToolTipText     =   "All Data will be displayed at once"
            Top             =   2400
            Width           =   1485
         End
      End
      Begin VB.Frame frameNavPanel 
         Caption         =   "Navigation Panel"
         Height          =   975
         Left            =   720
         TabIndex        =   25
         Top             =   6840
         Width           =   2295
         Begin VB.CommandButton cmdNEXT 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   29
            ToolTipText     =   "Next Record"
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton cmdLAST 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   28
            ToolTipText     =   "Jumps to Last Record"
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cmdFIRST 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   27
            ToolTipText     =   "Jumps to First Record"
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cmdPREV 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Sylfaen"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   26
            ToolTipText     =   "Previous Record"
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.Frame frameSearchBy 
         Caption         =   "SEARCH BY"
         Height          =   2055
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   3135
         Begin VB.CommandButton cmdSearchPrint 
            Caption         =   "Print "
            Height          =   375
            Left            =   1080
            TabIndex        =   24
            ToolTipText     =   "Selective Report will be generated based on search result"
            Top             =   1560
            Width           =   765
         End
         Begin VB.ComboBox cmbSearchDEPT_ID 
            Height          =   495
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Type here the value of Search"
            Top             =   960
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cmbSearchDEPT_NM 
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cmbSearchLOCATION 
            Height          =   495
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cmbSearchMANAGER 
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cmbSearchBUDGET 
            Height          =   495
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   375
            Left            =   1920
            TabIndex        =   18
            ToolTipText     =   "Search Specific Record"
            Top             =   1560
            Width           =   1005
         End
         Begin VB.ComboBox cmbSearchBy 
            Height          =   495
            Left            =   120
            TabIndex        =   17
            Text            =   " SEARCH BY"
            ToolTipText     =   "Search based on below Parameters"
            Top             =   480
            Width           =   2775
         End
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmDepartment.frx":94CA
      Height          =   4935
      Left            =   960
      TabIndex        =   14
      ToolTipText     =   "All Records are shown here"
      Top             =   3000
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      Caption         =   "DEPARTMENT TABLE"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "DEPT_ID"
         Caption         =   "DEPARTMENT ID"
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
         DataField       =   "DEPT_NM"
         Caption         =   "DEPARTMENT NAME"
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
      BeginProperty Column02 
         DataField       =   "LOCATION"
         Caption         =   "LOCATION"
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
      BeginProperty Column03 
         DataField       =   "MANAGER"
         Caption         =   "MANAGER"
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
      BeginProperty Column04 
         DataField       =   "BUDGET"
         Caption         =   "BUDGET"
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
         AllowRowSizing  =   -1  'True
         AllowSizing     =   -1  'True
         Size            =   3
         BeginProperty Column00 
            DividerStyle    =   3
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3600
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
   Begin VB.Frame Frame3 
      Caption         =   "DEPARTMENT ENTRY FORM"
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   12015
      Begin VB.ComboBox cmbDEPT_NM 
         Height          =   495
         Left            =   2640
         TabIndex        =   40
         Top             =   840
         Width           =   3135
      End
      Begin VB.ComboBox cmbManager 
         Height          =   495
         Left            =   7800
         TabIndex        =   39
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cmbLocation 
         Height          =   495
         Left            =   6000
         TabIndex        =   38
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtDEPT_ID 
         Height          =   495
         Left            =   480
         TabIndex        =   0
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtDEPT_NM 
         Height          =   495
         Left            =   2640
         TabIndex        =   1
         Text            =   "AAAAAAAAAAAAAAA"
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtLOCATION 
         Height          =   495
         Left            =   6000
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtMANAGER 
         Height          =   495
         Left            =   7800
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
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
         Left            =   9720
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblnum2txt 
         AutoSize        =   -1  'True
         Caption         =   "Rs Zero"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   9600
         TabIndex        =   41
         ToolTipText     =   "Number will be displayed as Text"
         Top             =   1440
         Width           =   540
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
         Left            =   4800
         TabIndex        =   13
         Top             =   360
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
         TabIndex        =   12
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department Id"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Department Name"
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   480
         Width           =   2220
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Manager"
         Height          =   375
         Left            =   7800
         TabIndex        =   8
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Budget"
         Height          =   375
         Left            =   9720
         TabIndex        =   7
         Top             =   480
         Width           =   840
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
      TabIndex        =   5
      Top             =   0
      Width           =   3210
   End
End
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Dim packet As String


Private Sub cmbDEPT_NM_GotFocus()
    'removes inside content
    cmbDEPT_NM.Clear
    
    'add content
    cmbDEPT_NM.AddItem "HEAD"
    cmbDEPT_NM.AddItem "ACCOUNTS"
    cmbDEPT_NM.AddItem "SALES"
    cmbDEPT_NM.AddItem "OPERATIONS"
    cmbDEPT_NM.AddItem "AddNEW"
End Sub

Private Sub cmbDEPT_NM_Click()
    If cmbDEPT_NM.Text = "AddNEW" Then
        cmbDEPT_NM.Visible = False
        txtDEPT_NM.Text = ""
    Else
        txtDEPT_NM.Text = cmbDEPT_NM.Text
    End If
End Sub


Private Sub cmbLocation_GotFocus()
    'removes inside content
    cmbLocation.Clear
    
    'add content
    cmbLocation.AddItem "FRNT01"
    cmbLocation.AddItem "FRNT02"
    cmbLocation.AddItem "BACK01"
    cmbLocation.AddItem "AddNew"
    
End Sub

Private Sub cmbLocation_Click()
    If cmbLocation.Text = "AddNew" Then
        cmbLocation.Visible = False
        txtLOCATION.Text = ""
    Else
        txtLOCATION.Text = cmbLocation.Text
    End If
End Sub

Private Sub cmbManager_GotFocus()
    'removes inside contents
    cmbManager.Clear
    
    'add contents
    cmbManager.AddItem "NMAGAR"
    cmbManager.AddItem "AJAYKR"
    cmbManager.AddItem "SNTOSH"
    cmbManager.AddItem "AddNew"
End Sub

Private Sub cmbManager_Click()
    If cmbManager.Text = "AddNew" Then
        cmbManager.Visible = False
        txtMANAGER.Text = ""
    Else
        txtMANAGER.Text = cmbManager.Text
    End If
End Sub

Private Sub cmbSearchBy_GotFocus()
    'removes inside contents
    cmbSearchBy.Clear
    
    'add contents
    cmbSearchBy.AddItem "Department ID"
    cmbSearchBy.AddItem "Department Name"
    cmbSearchBy.AddItem "Location"
    cmbSearchBy.AddItem "Manager"
    cmbSearchBy.AddItem "Budget"
End Sub

Private Sub cmbSearchBy_click()
    Dim response As String
    response = cmbSearchBy.Text

    'make visibility off
    
    cmbSearchDEPT_ID.Visible = False
    cmbSearchDEPT_NM.Visible = False
    cmbSearchLOCATION.Visible = False
    cmbSearchMANAGER.Visible = False
    cmbSearchBUDGET.Visible = False
    
    'validate response
    Select Case response
     Case "Department ID"
        cmbSearchDEPT_ID.Visible = True
        
        sql = "select distinct(dept_id) from department"
        cmd.CommandText = sql
        Set rcrdset = cmd.Execute
        
        'clear combobox
        cmbSearchDEPT_ID.Clear
        'record entry in combobox
        While Not rcrdset.EOF
            cmbSearchDEPT_ID.AddItem rcrdset.Fields(0) & ""
            rcrdset.MoveNext
        Wend
        
        cmbSearchDEPT_ID.SetFocus
     Case "Department Name"
        cmbSearchDEPT_NM.Visible = True
        
        sql = "select distinct(dept_nm) from department"
        cmd.CommandText = sql
        Set rcrdset = cmd.Execute
        
        'clear combobox
        cmbSearchDEPT_NM.Clear
        'record entry in combobox
        While Not rcrdset.EOF
            cmbSearchDEPT_NM.AddItem rcrdset.Fields(0) & ""
            rcrdset.MoveNext
        Wend
        
        cmbSearchDEPT_NM.SetFocus
     Case "Location"
        cmbSearchLOCATION.Visible = True
        
        sql = "select distinct(location) from department"
        cmd.CommandText = sql
        Set rcrdset = cmd.Execute
        
        'clear combobox
        cmbSearchLOCATION.Clear
        'record entry in combobox
        While Not rcrdset.EOF
            cmbSearchLOCATION.AddItem rcrdset.Fields(0) & ""
            rcrdset.MoveNext
        Wend
        
        cmbSearchLOCATION.SetFocus
        
     Case "Manager"
        cmbSearchMANAGER.Visible = True
        
        sql = "select distinct(manager) from department"
        cmd.CommandText = sql
        Set rcrdset = cmd.Execute
        
        'clear combobox
        cmbSearchMANAGER.Clear
        'record entry in combobox
        While Not rcrdset.EOF
            cmbSearchMANAGER.AddItem rcrdset.Fields(0) & ""
            rcrdset.MoveNext
        Wend
        
        cmbSearchMANAGER.SetFocus
        
     Case "Budget"
        cmbSearchBUDGET.Visible = True
        
        sql = "select distinct(budget) from department"
        cmd.CommandText = sql
        Set rcrdset = cmd.Execute
        
        'clear combobox
        cmbSearchBUDGET.Clear
        'record entry in combobox
        While Not rcrdset.EOF
            cmbSearchBUDGET.AddItem rcrdset.Fields(0) & ""
            rcrdset.MoveNext
        Wend
        
        cmbSearchBUDGET.SetFocus
    End Select

End Sub



Private Sub cmbOFF()
    cmbSearchDEPT_ID.Visible = flag
    cmbSearchDEPT_NM.Visible = flag
    cmbSearchLOCATION.Visible = flag
    cmbSearchMANAGER.Visible = flag
    cmbSearchBUDGET.Visible = flag
End Sub





Private Sub cmdCREATE_Click()

  'first job: disable buttons
    cmdCREATE.Enabled = False
    cmdUPDATE.Enabled = False
    cmdDELETE.Enabled = False
    cmdSAVE.Enabled = True
    cmdEXIT.Enabled = True
    cmdPrintAll.Enabled = False
    cmdDisplay.Enabled = False
    
   'set focus
    cmdSAVE.SetFocus
    
  'second job: primary key set
    
    'fetching department data to check if record exists or not
    sql = "SELECT * from department"
    cmd.CommandText = sql
    Set rcrdset = cmd.Execute
    
    'check if primary should reset
    If rcrdset.EOF And rcrdset.BOF Then
        'pk automatic generation resets
        sql = "update dept_id set pk = 0"
        cmd.CommandText = sql
        cmd.Execute
    End If
    
    'primary key value generation
    sql = "SELECT * from dept_id"
    cmd.CommandText = sql
    Set rcrdset = cmd.Execute
    txtDEPT_ID.Text = "DEPT" & (rcrdset.Fields(0).Value + 1)


 'It will be shared inside cmdSave CodeBlock
    packet = "create"
End Sub
Private Sub cmdUPDATE_Click()
    
    'appear combobox
    cmbDEPT_NM.Visible = True
    cmbLocation.Visible = True
    cmbManager.Visible = True
    
  'Enable/Disable buttons
    cmdCREATE.Enabled = False
    cmdUPDATE.Enabled = False
    cmdDELETE.Enabled = False
    cmdSAVE.Enabled = True
    cmdEXIT.Enabled = True
    cmdDisplay.Enabled = False
    cmdPrintAll.Enabled = False
    
 'set focus
    cmdSAVE.SetFocus
 'inform user
    MsgBox "UPDATE CLICKED !!!." & vbCrLf & "Please change or update the fields as needed" & vbCrLf & "and then press 'Save'.", vbInformation + vbOKOnly, "Update Operation"
    
 'It will be shared inside cmdSave CodeBlock
    packet = "update"
End Sub

Private Sub cmdDELETE_Click()
    
  'Enable/Disable buttons
    cmdCREATE.Enabled = False
    cmdUPDATE.Enabled = False
    cmdDELETE.Enabled = False
    cmdSAVE.Enabled = True
    cmdEXIT.Enabled = True
    cmdDisplay.Enabled = False
    cmdPrintAll.Enabled = False

    'set focus
    cmdSAVE.SetFocus
    'It will be shared inside cmdSave CodeBlock
    packet = "delete"
End Sub

Private Sub cmdSAVE_Click()
    
    ' Ask for confirmation before saving
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to save this record?", vbQuestion + vbYesNo, "Confirm Save")
    If response = vbNo Then
        MsgBox "Operation cancelled.", vbInformation + vbOKOnly, "Cancelled"
        Exit Sub
    End If
    If response = vbYes Then
        Select Case packet
            Case "create"
                sql = "insert into department values(?,?,?,?,?)"
                With cmd
                    .CommandText = sql
                    .Parameters.Append .CreateParameter("dept_id", adVarChar, adParamInput, 6, txtDEPT_ID.Text)
                    .Parameters.Append .CreateParameter("dept_nm", adVarChar, adParamInput, 15, txtDEPT_NM.Text)
                    .Parameters.Append .CreateParameter("location", adVarChar, adParamInput, 6, txtLOCATION.Text)
                    .Parameters.Append .CreateParameter("manager", adVarChar, adParamInput, 6, txtMANAGER.Text)
                    .Parameters.Append .CreateParameter("budget", adVarChar, adParamInput, 10, txtBUDGET.Text)
                End With
                Set rcrdset = cmd.Execute
                
                'UPDATING SEQUENCE
                sql = "UPDATE DEPT_ID SET PK = PK +1"
                cmd.CommandText = sql
    
            Case "update"
                ' Set up the SQL command and parameters
                sql = "update department set dept_id = ?, dept_nm = ?, location = ?, manager = ?, budget = ? where dept_id = ?"
                With cmd
                    .CommandText = sql
                    .Parameters.Append .CreateParameter("dept_id", adVarChar, adParamInput, 6, txtDEPT_ID.Text)
                    .Parameters.Append .CreateParameter("dept_nm", adVarChar, adParamInput, 15, txtDEPT_NM.Text)
                    .Parameters.Append .CreateParameter("location", adVarChar, adParamInput, 6, txtLOCATION.Text)
                    .Parameters.Append .CreateParameter("manager", adVarChar, adParamInput, 6, txtMANAGER.Text)
                    .Parameters.Append .CreateParameter("budget", adVarChar, adParamInput, 10, txtBUDGET.Text)
                    .Parameters.Append .CreateParameter("dept_id", adVarChar, adParamInput, 6, txtDEPT_ID.Text) ' For the WHERE clause
                End With
            
            Case "delete"
                sql = "delete from department where dept_id = ?"
                With cmd
                    .CommandText = sql
                    .Parameters.Append .CreateParameter("dept_id", adVarChar, adParamInput, 6, txtDEPT_ID.Text)
                End With
    
            End Select
    'execute command
    Set rcrdset = cmd.Execute
    
    'inform user
    response = MsgBox("Operation completed successfully!", vbInformation + vbOKOnly, "Success")
    
    'first job: enable/disable buttons
    cmdCREATE.Enabled = True
    cmdUPDATE.Enabled = False
    cmdDELETE.Enabled = False
    cmdSAVE.Enabled = True
    cmdEXIT.Enabled = True
    cmdPrintAll.Enabled = True
    cmdDisplay.Enabled = True
    
    ' Clear existing parameters
        Dim i As Integer
        For i = cmd.Parameters.Count - 1 To 0 Step -1
          cmd.Parameters.Delete i
        Next i
    End If
    
    'Refresh datagrid
    Adodc1.Refresh
    
    'clear input fields
    txtDEPT_ID.Text = ""
    txtDEPT_NM.Text = ""
    txtLOCATION.Text = ""
    txtMANAGER.Text = ""
    txtBUDGET.Text = ""
       
    cmbDEPT_NM.Text = ""
    cmbLocation.Text = ""
    cmbManager.Text = ""
  
End Sub

Private Sub cmdSearch_Click()
'sql = "select * from department where dept_id = '" + txtDEPT_ID.Text + "' or dept_nm = '" + txtDEPT_NM.Text + "' or  location = '" + txtLOCATION.Text + "' or manager = '" + txtMANAGER.Text + "' or budget = '" + txtBUDGET.Text + "'"
sql = "select * from department where dept_id = ? or dept_nm = ? or location = ? or manager = ? or budget = ?"
   
With cmd
    .CommandText = sql
    .Parameters.Append .CreateParameter("dept_id", adVarChar, adParamInput, 6, cmbSearchDEPT_ID.Text)
    .Parameters.Append .CreateParameter("dept_nm", adVarChar, adParamInput, 15, cmbSearchDEPT_NM.Text)
    .Parameters.Append .CreateParameter("location", adVarChar, adParamInput, 6, cmbSearchLOCATION.Text)
    .Parameters.Append .CreateParameter("manager", adVarChar, adParamInput, 6, cmbSearchMANAGER.Text)
    .Parameters.Append .CreateParameter("budget", adVarChar, adParamInput, 10, cmbSearchBUDGET.Text)
End With

    Set rcrdset = cmd.Execute

'If rcrdset.EOF Then
'   response = MsgBox("Record Not Found", vbInformation, "Message")
'    Exit Sub
'End If
    
'data display in entry form
txtDEPT_ID.Text = rcrdset.Fields(0).Value & ""
txtDEPT_NM.Text = rcrdset.Fields(1).Value & ""
txtLOCATION.Text = rcrdset.Fields(2).Value & ""
txtMANAGER.Text = rcrdset.Fields(3).Value & ""
txtBUDGET.Text = rcrdset.Fields(4).Value & ""

cmbDEPT_NM.Text = rcrdset.Fields(1).Value & ""
cmbLocation.Text = rcrdset.Fields(2).Value & ""
cmbManager.Text = rcrdset.Fields(3).Value & ""

  'Enable/Disable buttons
    cmdCREATE.Enabled = True
    cmdUPDATE.Enabled = True
    cmdDELETE.Enabled = True
    cmdSAVE.Enabled = False
    cmdEXIT.Enabled = True
    cmdDisplay.Enabled = True
    cmdPrintAll.Enabled = True

' Clear existing parameters
Dim i As Integer
For i = cmd.Parameters.Count - 1 To 0 Step -1
    cmd.Parameters.Delete i
Next i
End Sub

Private Sub cmdSearchPrint_Click()
    ' Set the parameter for Command2
    
    DataEnvironment1.Commands("Command2").Parameters(0).Value = cmbSearchDEPT_ID.Text
    DataEnvironment1.Commands("Command2").Parameters(1).Value = cmbSearchDEPT_NM.Text
    DataEnvironment1.Commands("Command2").Parameters(2).Value = cmbSearchLOCATION.Text
    DataEnvironment1.Commands("Command2").Parameters(3).Value = cmbSearchMANAGER.Text
    DataEnvironment1.Commands("Command2").Parameters(4).Value = cmbSearchBUDGET.Text
    
    selectiveReport.Show
End Sub
Private Sub cmdDisplay_Click()
Adodc1.Refresh
If Adodc1.Recordset.EOF And Adodc1.Recordset.BOF Then
    MsgBox "No records available." & vbCrLf & "Please create a record first then try again later.", vbExclamation + vbOKOnly, "No Records Found"
    Exit Sub
End If

flag = Not flag
DataGrid1.Visible = flag
End Sub

Private Sub cmdPrintAll_Click()
If Adodc1.Recordset.EOF And Adodc1.Recordset.BOF Then
    MsgBox "No records available." & vbCrLf & "Please create a record first then try again later.", vbExclamation + vbOKOnly, "No Records Found"
    Exit Sub
End If

collectiveReport.Show
End Sub



Private Sub cmdEXIT_Click()
    ' Display a confirmation message box
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you want to exit?", vbYesNoCancel + vbQuestion, "Exit Confirmation")
    
    ' Handle the user's response
    Select Case response
        Case vbYes
            ' Unload the form to close it
            Unload Me
        Case vbNo
            ' Display a message and return to the application
            MsgBox "You chose to continue using the application.", vbInformation, "Continue"
        Case vbCancel
            ' Display a message and return to the application
            MsgBox "Action cancelled.", vbInformation, "Cancelled"
    End Select
End Sub


Private Sub cmdFIRST_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdLAST_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdNEXT_Click()
    If Not Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveNext
    Else
        Adodc1.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdPREV_Click()
    If Not Adodc1.Recordset.BOF Then
        Adodc1.Recordset.MovePrevious
    Else
        Adodc1.Recordset.MoveLast
    End If
End Sub






Private Sub Form_Load()
    flag = False
'////////////////////////////////////////////////////////////////////////
    'creating connection object
    Set conn = New ADODB.Connection
    'setting connection String
    connString = "Provider=MSDAORA.1;User ID=PRJ2432N/PRJ2432N;Persist Security Info=False"
    'connection open
    conn.Open connString
    
    Set cmd = New ADODB.Command
    'parameterised query for
    '1.Query Caching(faster execution at Cache memory)
    '2.protection from SQL Injections
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        ' Add the parameter
        '.Parameters.Append .CreateParameter("MyParam", adVarChar, adParamInput, 50, paramValue)
    End With
    Set rcrdset = New ADODB.Recordset
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub txtBUDGET_Change()
    Dim number As Long
    number = Val(txtBUDGET.Text)
    lblnum2txt.Caption = "Rs " & ConvertNumberToText(number)
End Sub
