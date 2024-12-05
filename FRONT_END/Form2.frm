VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   7620
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim strConn As String
Dim strPLSQL As String



Private Sub Command1_Click()
conn.Open strConn

MsgBox ("HI")
End Sub

Private Sub Form_Load()
' Initialize objects
Set conn = New ADODB.Connection
Set cmd = New ADODB.Command
Set rs = New ADODB.Recordset

' Connection string for Oracle (update with your database details)
strConn = "Provider=MSDAORA.1;User ID=PRJ2432N/PRJ2432N;Persist Security Info=False"
Print "CONNECTION DONE"
End Sub





