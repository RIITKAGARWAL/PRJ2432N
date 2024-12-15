VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Department & Position"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   6750
   ScaleWidth      =   13830
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Line Line2 
      X1              =   8040
      X2              =   8040
      Y1              =   0
      Y2              =   8640
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   16080
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    responsive Command1
End Sub

Private Sub Form_Resize()
    responsive Command1, 3480, 2760, 975, 3975
    responsive Command2, 3480, 2760, 975, 3975
End Sub

Private Sub responsive(ctrl As CommandButton, left As Integer, _
top As Integer, height As Integer, width As Integer)
    ctrl.left = left * (Me.width / 16395)
    ctrl.top = top * (Me.height / 9270)
    ctrl.width = width * (Me.width / 16395)
    ctrl.height = height * (Me.height / 9270)
End Sub



