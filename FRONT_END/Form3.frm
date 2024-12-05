VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{D5288401-E6C5-11D1-BE7D-C63815000000}#1.0#0"; "CHARTWIZ.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   LinkTopic       =   "Form3"
   ScaleHeight     =   7545
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin MSChartWiz.SubWizard SubWizard1 
      Height          =   1935
      Left            =   2160
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3413
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   5640
      Top             =   360
      _ExtentX        =   10583
      _ExtentY        =   7938
      _Version        =   393216
      Picture         =   "Form3.frx":0000
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   720
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4215
      Left            =   5520
      TabIndex        =   1
      Top             =   1800
      Width           =   5535
      ExtentX         =   9763
      ExtentY         =   7435
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
      Location        =   "http:///"
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8520
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   5
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   443
      URL             =   "https://127.0.0.1"
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3615
      Left            =   600
      OleObjectBlob   =   "Form3.frx":1D912
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Scriptlet1_onscriptletevent(ByVal name As String, ByVal eventData As Variant)

End Sub

