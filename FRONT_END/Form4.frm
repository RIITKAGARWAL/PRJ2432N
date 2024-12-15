VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16575
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   16575
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "ms edge open"
      Height          =   615
      Left            =   5640
      TabIndex        =   2
      Top             =   7560
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "search"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   7560
      Width           =   4095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      ExtentX         =   23733
      ExtentY         =   12515
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WebBrowser1.navigate "https://lets-vibe-with-music.netlify.app/"
End Sub

Private Sub Command2_Click()
    ' Open Microsoft Edge and navigate to a specific URL
    Dim result As Long
    Dim url As String
    url = "https://www.example.com"
    result = Shell("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & url, vbNormalFocus)
End Sub

