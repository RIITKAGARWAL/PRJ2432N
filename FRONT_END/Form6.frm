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
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   16155
   WindowState     =   2  'Maximized
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
      TabIndex        =   18
      Top             =   720
      Width           =   5655
      Begin VB.TextBox txtDEPT_ID 
         Enabled         =   0   'False
         Height          =   495
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtDEPT_NM 
         Height          =   495
         Left            =   3000
         TabIndex        =   1
         Top             =   990
         Width           =   2415
      End
      Begin VB.TextBox txtLOCATION 
         Height          =   495
         Left            =   3000
         TabIndex        =   2
         Top             =   1605
         Width           =   2415
      End
      Begin VB.TextBox txtMANAGER 
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   2235
         Width           =   2415
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
         Left            =   3000
         TabIndex        =   4
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Department Name"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Manager"
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Budget"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   2880
         Width           =   840
      End
   End
   Begin MSDataGridLib.DataGrid DataGridDept 
      Bindings        =   "Form6.frx":94CA
      Height          =   2535
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Appearance      =   0
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
      TabIndex        =   11
      Top             =   4440
      Width           =   5655
      Begin VB.CommandButton cmdPREV 
         Caption         =   "&Previous"
         Height          =   615
         Left            =   1485
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdFIRST 
         Caption         =   "&First"
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdLAST 
         Caption         =   "&Last"
         Height          =   615
         Left            =   4320
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdNEXT 
         Caption         =   "&Next"
         Height          =   615
         Left            =   3195
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control Panel"
      Height          =   1335
      Left            =   10440
      TabIndex        =   5
      Top             =   4440
      Width           =   5655
      Begin VB.CommandButton cmdEXIT 
         Caption         =   "&Exit"
         Height          =   615
         Left            =   4560
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdUPDATE 
         Caption         =   "&Update"
         Height          =   615
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdDELETE 
         Caption         =   "&Delete"
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdCREATE 
         Caption         =   "&Create"
         Height          =   615
         Left            =   120
         TabIndex        =   6
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
      TabIndex        =   17
      Top             =   0
      Width           =   3210
   End
End
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+++++++++Global Declaration Zone++++++
Dim toogle As Boolean
Dim packet As String
'----------Global Declaration Zone Ends here ---------



'+++++++form load code begins here+++++++
Private Sub Form_Load()
    
    'lock all textbox
    TextBoxes "locked", True
    
    'initializing toogle
    toogle = True
    
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
    'subroutine called
    navPnl
End Sub
'----------form load code ends here-------
Private Sub BofEof()
If ((rcrdset.BOF = True) Or (rcrdset.EOF = True)) Then
MsgBox ("No Records Available")
End If
End Sub


'+++++++++ NavPanel code begins here++++++++

Private Sub cmdFIRST_Click()
   On Error GoTo NAV_ERR_HNDLR
   rcrdset.MoveFirst
Exit Sub
NAV_ERR_HNDLR:
    BofEof
End Sub

Private Sub cmdLAST_Click()
On Error GoTo NAV_ERR_HNDLR
   rcrdset.MoveLast
Exit Sub
NAV_ERR_HNDLR:
    BofEof
End Sub

Private Sub cmdPREV_Click()
On Error GoTo NAV_ERR_HNDLR
    If Not rcrdset.BOF Then
        rcrdset.MovePrevious
    Else
        rcrdset.MoveLast
    End If
Exit Sub
NAV_ERR_HNDLR:
    BofEof
End Sub
Private Sub cmdNEXT_Click()
On Error GoTo NAV_ERR_HNDLR
    If Not rcrdset.EOF Then
        rcrdset.MoveNext
    Else
        rcrdset.MoveFirst
    End If
Exit Sub
NAV_ERR_HNDLR:
    BofEof
End Sub
'----------Nav Panel code ends here----------

'+++++++++++ control panel block begins here ++++++

Private Sub cmdCREATE_Click()
    'set focus on txtDEPT_NM
    txtDEPT_NM.SetFocus
    'unlock all textbox
    TextBoxes "locked", False
    'clear all Textbox texts
    TextBoxes "blank", False
    
    'disable all button
    cmdOnOff False
    
    'enable Save Button
    cmdSAVE.Enabled = True
    
    On Error GoTo BCKEND_ERR_HNDLR

    'fetching sequence from oracle 11g
    sql = "select PK + 1 FROM DEPT_ID"
    
    cmd.CommandText = sql
    Set rcrdset = cmd.Execute
    
    'making primary key alpha-numeric
    Dim PK As String
    PK = "DEPT" & (rcrdset.Fields(0).Value)
    txtDEPT_ID.Text = PK
    
    'It will be shared inside cmdSave CodeBlock
    packet = "create"
Exit Sub

'catch block of bckendErrHndlr
BCKEND_ERR_HNDLR:
    'sub procedure from module2
    ErrHndlCode
End Sub

Private Sub cmdUPDATE_Click()

    'set focus on txtDEPT_NM
    txtDEPT_NM.SetFocus
    
    'unlock all textbox
    TextBoxes "locked", False
  
    'disable all button
    cmdOnOff False
    
    'enable Save Button
    cmdSAVE.Enabled = True
    
    'It will be shared inside cmdSave CodeBlock
    packet = "update"

End Sub

Private Sub cmdDELETE_Click()
    'set focus on cmdDelete
    cmdDELETE.SetFocus
    'lock all textbox
    TextBoxes "locked", True
    
    'disable all button
    cmdOnOff False
    
    'enable Save Button
    cmdSAVE.Enabled = True
    
    packet = "delete"
    
End Sub

Private Sub cmdSAVE_Click()
On Error GoTo BCKEND_ERR_HNDLR
    cmdOnOff False
    Dim response As Integer
    'validate from user
    response = MsgBox("Do you want to save changes?", vbYesNo, "Save Changes")

    If response = vbYes Then
    
        Select Case packet
            Case "create"
                'SQL statement
                sql = "insert into department values('" + txtDEPT_ID.Text + "','" + txtDEPT_NM.Text + "','" + txtLOCATION + "','" + txtMANAGER + "','" + txtBUDGET + "')"
                cmd.CommandText = sql
                Set rcrdset = cmd.Execute
                
                'UPDATING SEQUENCE
                sql = "UPDATE DEPT_ID SET PK = PK +1"
                cmd.CommandText = sql
                Set rcrdset = cmd.Execute
                
            Case "update"
                'SQL statement
                sql = "update department set dept_nm = '" + txtDEPT_NM.Text + "',location = '" + txtLOCATION + "',manager = '" + txtMANAGER + "',budget = '" + txtBUDGET + "' where dept_id = '" + txtDEPT_ID.Text + "' "
                cmd.CommandText = sql
                Set rcrdset = cmd.Execute
                
            Case "delete"
                'SQL statement
                sql = "delete from department where dept_id =  '" + txtDEPT_ID.Text + "'"
                cmd.CommandText = sql
                Set rcrdset = cmd.Execute
        End Select
        'reset the string
        packet = ""
       
        response = MsgBox("Record Saved!", vbInformation)
    Else
        response = MsgBox("Record creation Failed!", vbCritical)
    End If
    Adodc1.Refresh
    TextBoxes "blank", False
    TextBoxes "locked", True
    
    cmdOnOff True
    cmdSAVE.Enabled = False
    navPnl
Exit Sub

'catch block of bckendErrHndlr
BCKEND_ERR_HNDLR:

    'reset the string
    packet = ""
    'calling sub routine
    ErrHndlCode
    
    'SQL statement
    sql = "select * from department"
    cmd.CommandText = sql
    Set rcrdset = cmd.Execute
    
    'after error handling task to do
    TextBoxes "blank", False
    TextBoxes "locked", True
    navPnl
    cmdOnOff True
    cmdSAVE.Enabled = False
End Sub

Private Sub cmdEXIT_Click()
    Dim response As Integer
    response = MsgBox("Do you want to Exit?", vbYesNoCancel)
    
    Select Case response
        Case vbYes
            End
            ' Code to abort
        Case vbNo
            ' Code to retry
        Case vbCancel
            ' Code to ignore
    End Select
End Sub
'-------control panel code ends here-------

'+++++++++display panel block begins here++++++
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
'------display panel block ends here------

'+++++ procedures below ++++++

Private Sub cmdOnOff(response As Boolean)
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is CommandButton Then
         ctrl.Enabled = response
        End If
    Next ctrl
End Sub
Private Sub TextBoxes(packet As String, state As Boolean)
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            Select Case packet
                Case "blank"
                    ctrl.Text = ""
                Case "locked"
                    ctrl.Locked = state
            End Select
        End If
    Next ctrl
End Sub

Private Sub navPnl()
    If rcrdset.state = adStateOpen Then
        rcrdset.Close
    End If
    rcrdset.Open "SELECT * FROM department", conn, adOpenKeyset, adLockOptimistic

    ' Check if the recordset is empty
    If rcrdset.EOF And rcrdset.BOF Then
        cmdOnOff False
        cmdCREATE.Enabled = True
        TextBoxes "blank", False
        TextBoxes "locked", True
        Exit Sub
    End If

    ' Bind the recordset to the form controls
    Set txtDEPT_ID.DataSource = rcrdset
    txtDEPT_ID.DataField = "DEPT_ID"
    
    Set txtDEPT_NM.DataSource = rcrdset
    txtDEPT_NM.DataField = "DEPT_NM"
    
    Set txtLOCATION.DataSource = rcrdset
    txtLOCATION.DataField = "LOCATION"
    
    Set txtMANAGER.DataSource = rcrdset
    txtMANAGER.DataField = "MANAGER"
    
    Set txtBUDGET.DataSource = rcrdset
    txtBUDGET.DataField = "BUDGET"
End Sub

'--------------------placeholder-------------
'----------------tab & enter coding-----------

'----------------properties click-------------

'-------------------responsiveness testing---------------------

'----------------Backend Integration--------------
Private Sub bckendIntegration()
    'raise error if problem occurs during execution
    On Error GoTo BCKEND_ERR_HNDLR
    
    'creating a connection object
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    
    'declaring and assigning connection String
    Dim connString As String
    connString = "Provider=MSDAORA.1;User ID=PRJ2432N/PRJ2432N;Persist Security Info=False"
    
    'opening connection string
    conn.Open connString
    
    'creating the command object
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    'parameterised query for
    '1.Query Caching(faster execution at Cache memory)
    '2.protection from SQL Injections
    With cmd
        .ActiveConnection = conn
        .CommandText = "Select * from DEPARTMENT"
        .CommandType = adCmdText
        ' Add the parameter
        '.Parameters.Append .CreateParameter("MyParam", adVarChar, adParamInput, 50, paramValue)
    End With
    
    'creating the Recordset object
    Dim rcrdset As ADODB.Recordset
    Set rcrdset = cmd.Execute
    
    
    Label1.Caption = rcrdset.Fields(0) & ""
    Text1.Text = rcrdset.Fields(1) & ""
    Text2.Text = rcrdset.Fields(2) & ""
    Text3.Text = rcrdset.Fields(3) & ""
    
    'closing the recordset
    rcrdset.Close
    'freeing up the container
    Set rcrdset = Nothing
    
    'closing the connection
    conn.Close
    'freeing up the container
    Set conn = Nothing
    
    'Everything ran smoothly, Success!
    Exit Sub
    
'If runtime error occurs in any part of bckend sub
'the compiler will jump directly to the below part of code

'catch block of bckendErrHndlr
BCKEND_ERR_HNDLR:
    'Informing User about the Runtime Error
    MsgBox ("Something Unexpected has Occured! " & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.Description & Chr(13) & Chr(10) & _
        "Error Source: " & Err.Source)
    
    'clear the Error
    Err.Clear
    
    If Not conn Is Nothing Then
        'closing the connection
        conn.Close
        'freeing up the container
        Set conn = Nothing
    End If
    'Error Successfully Dealt!
End Sub


Private Sub txtDEPT_NM_KeyPress(keyascii As Integer)
    enterKeyPress txtDEPT_NM, txtLOCATION, keyascii
    isValidInput txtDEPT_NM, "alphabet", keyascii, 0
    isValidInput txtDEPT_NM, "length", keyascii, 15
    isValidInput txtDEPT_NM, "empty", keyascii, 0
End Sub



Private Sub txtDEPT_NM_LostFocus()

                If Trim(txtDEPT_NM.Text) = "" Then
                    txtDEPT_NM.SetFocus
                    response = MsgBox("The Field is Mandatory", vbExclamation)
                End If
                
End Sub

Private Sub txtLOCATION_KeyPress(keyascii As Integer)
    enterKeyPress txtLOCATION, txtMANAGER, keyascii
     isValidInput txtLOCATION, "length", keyascii, 6
End Sub

Private Sub txtMANAGER_KeyPress(keyascii As Integer)
    enterKeyPress txtMANAGER, txtBUDGET, keyascii
    isValidInput txtMANAGER, "length", keyascii, 6
End Sub

Private Sub txtBUDGET_KeyPress(keyascii As Integer)
   ' enterKeyPress txtBUDGET, txtBUDGET, KeyAscii
   isValidInput txtBUDGET, "decimal", keyascii, 0
   If Val(txtBUDGET.Text) < 999999 And keyascii = vbKeyBack Then
    keyascii = 0 'acts as a backspace
   End If
End Sub

