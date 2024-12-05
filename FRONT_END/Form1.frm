VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   8670
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdb 
      Left            =   7560
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LOAD SQL FILE"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   3135
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXECUTE PL/SQL"
      Height          =   855
      Left            =   3960
      TabIndex        =   0
      Top             =   3960
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filePath As String
Dim C As ADODB.Connection
Dim connString As String
Dim sql As String

Private Sub Form_Load()
    connString = "Provider=MSDAORA.1;User ID=PRJ2432N/PRJ2432N;Persist Security Info=False"
End Sub

Private Sub Command1_Click()
On Error GoTo ERRHANDLE
    Set C = New ADODB.Connection
    C.Open connString
    sql = RichTextBox1.Text
    sql = Replace(sql, vbCrLf, "")
    C.Execute sql
    MsgBox "PL/SQL executed successfully!"
    C.Close
    Set C = Nothing
    Exit Sub
ERRHANDLE:
    MsgBox "SOMETHING HAPPENED"
End Sub

Private Sub Command2_Click()
    cdb.Filter = "All Files (*.*)|*.*"
    On Error GoTo ERRORHANDLING
    cdb.ShowOpen
    filePath = cdb.FileName
    If filePath <> "" Then
        RichTextBox1.LoadFile filePath, rtfText
    End If
    Exit Sub
ERRORHANDLING:
    MsgBox "Operation cancelled."
End Sub


