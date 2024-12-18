VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmp 
   Caption         =   "Employee"
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
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   16155
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame frmDispPnl 
      Caption         =   "Display Panel"
      Height          =   1215
      Left            =   10800
      TabIndex        =   135
      Top             =   4920
      Width           =   4935
      Begin VB.CommandButton cmdEmpDisp 
         Caption         =   "&Display Records"
         Height          =   615
         Left            =   120
         TabIndex        =   137
         Top             =   480
         Width           =   2200
      End
      Begin VB.CommandButton cmdEmpSearch 
         Caption         =   "Search &Panel"
         Height          =   615
         Left            =   2520
         TabIndex        =   136
         Top             =   480
         Width           =   2200
      End
   End
   Begin VB.Frame frmSearch 
      Caption         =   "SEARCH BY"
      Height          =   4575
      Left            =   120
      TabIndex        =   105
      Top             =   0
      Visible         =   0   'False
      Width           =   15855
      Begin VB.ComboBox Combo18 
         Height          =   495
         Left            =   10680
         TabIndex        =   133
         Text            =   " BasicPay"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List5 
         Height          =   1275
         Left            =   10680
         Style           =   1  'Checkbox
         TabIndex        =   132
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo17 
         Height          =   495
         Left            =   9240
         TabIndex        =   131
         Text            =   " Level"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox Combo16 
         Height          =   495
         Left            =   2520
         TabIndex        =   130
         Text            =   " PositionID"
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox Combo15 
         Height          =   495
         Left            =   6600
         TabIndex        =   129
         Text            =   " EmployeeName"
         Top             =   480
         Width           =   2535
      End
      Begin VB.ListBox List2 
         Height          =   1275
         Left            =   6600
         Style           =   1  'Checkbox
         TabIndex        =   128
         Top             =   960
         Width           =   2535
      End
      Begin VB.ListBox List3 
         Height          =   1275
         Left            =   2520
         Style           =   1  'Checkbox
         TabIndex        =   127
         Top             =   960
         Width           =   1935
      End
      Begin VB.ListBox List4 
         Height          =   1275
         Left            =   9240
         Style           =   1  'Checkbox
         TabIndex        =   126
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox Combo7 
         Height          =   495
         Left            =   120
         TabIndex        =   125
         Text            =   " DepartmentID"
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Height          =   1275
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   124
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox Combo4 
         Height          =   495
         Left            =   4560
         TabIndex        =   123
         Text            =   " EmployeeId"
         Top             =   480
         Width           =   1935
      End
      Begin VB.ListBox List6 
         Height          =   1275
         Left            =   4560
         Style           =   1  'Checkbox
         TabIndex        =   122
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox Combo6 
         Height          =   495
         Left            =   3000
         TabIndex        =   121
         Text            =   "Empty"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ListBox List7 
         Height          =   1275
         Left            =   3000
         Style           =   1  'Checkbox
         TabIndex        =   120
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   495
         Left            =   1560
         TabIndex        =   119
         Text            =   " Gender"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ListBox List8 
         Height          =   1275
         Left            =   1560
         Style           =   1  'Checkbox
         TabIndex        =   118
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox Combo9 
         Height          =   495
         Left            =   13800
         TabIndex        =   117
         Text            =   " Status"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox List9 
         Height          =   1275
         Left            =   13800
         Style           =   1  'Checkbox
         TabIndex        =   116
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo10 
         Height          =   495
         Left            =   12360
         TabIndex        =   115
         Text            =   " Date of Joining"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ListBox List10 
         Height          =   1275
         Left            =   12360
         Style           =   1  'Checkbox
         TabIndex        =   114
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox Combo11 
         Height          =   495
         Left            =   8400
         TabIndex        =   113
         Text            =   " PAN"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ListBox List11 
         Height          =   1275
         Left            =   8400
         Style           =   1  'Checkbox
         TabIndex        =   112
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox Combo12 
         Height          =   495
         Left            =   6360
         TabIndex        =   111
         Text            =   " Aadhar"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ListBox List12 
         Height          =   1275
         Left            =   6360
         Style           =   1  'Checkbox
         TabIndex        =   110
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ComboBox Combo13 
         Height          =   495
         Left            =   4680
         TabIndex        =   109
         Text            =   " Type"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ListBox List13 
         Height          =   1275
         Left            =   4680
         Style           =   1  'Checkbox
         TabIndex        =   108
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox Combo14 
         Height          =   495
         Left            =   120
         TabIndex        =   107
         Text            =   " Age"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ListBox List14 
         Height          =   1275
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   106
         Top             =   3000
         Width           =   1335
      End
   End
   Begin VB.Frame frmEmpEntry 
      Caption         =   "EMPLOYEE ENTRY FORM"
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   15735
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   8640
         TabIndex        =   96
         Top             =   4920
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   6480
         TabIndex        =   95
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Frame frmNavPnl 
         Caption         =   "Navigation Panel"
         Height          =   1215
         Left            =   8280
         TabIndex        =   82
         Top             =   5520
         Width           =   7335
         Begin VB.CommandButton Command10 
            Caption         =   "&Previous"
            Height          =   615
            Left            =   1965
            TabIndex        =   86
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton Command9 
            Caption         =   "&First"
            Height          =   615
            Left            =   240
            TabIndex        =   85
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton Command8 
            Caption         =   "&Last"
            Height          =   615
            Left            =   5400
            TabIndex        =   84
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton Command7 
            Caption         =   "&Next"
            Height          =   615
            Left            =   3675
            TabIndex        =   83
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame frmCtrlPnl 
         Caption         =   "Control Panel"
         Height          =   1215
         Left            =   8280
         TabIndex        =   76
         Top             =   6720
         Width           =   7335
         Begin VB.CommandButton Command5 
            Caption         =   "&Exit"
            Height          =   615
            Left            =   5880
            TabIndex        =   81
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Update"
            Height          =   615
            Left            =   1560
            TabIndex        =   80
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Save"
            Height          =   615
            Left            =   4440
            TabIndex        =   79
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Delete"
            Height          =   615
            Left            =   3000
            TabIndex        =   78
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Create"
            Height          =   615
            Left            =   120
            TabIndex        =   77
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   495
         Left            =   13440
         TabIndex        =   74
         Text            =   " Type"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   495
         Left            =   13440
         TabIndex        =   73
         Text            =   " Status"
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Frame frmCntctDetal 
         Caption         =   "Contact Details"
         Height          =   3255
         Left            =   360
         TabIndex        =   60
         Top             =   4560
         Width           =   5895
         Begin VB.TextBox Text17 
            Height          =   495
            Left            =   3360
            TabIndex        =   65
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox Text18 
            Height          =   495
            Left            =   3360
            TabIndex        =   64
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox Text19 
            Height          =   495
            Left            =   1440
            TabIndex        =   63
            Top             =   2640
            Width           =   4335
         End
         Begin VB.TextBox Text20 
            Height          =   495
            Left            =   3360
            TabIndex        =   62
            Top             =   2040
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Whatsapp Number same as Phone Number"
            Height          =   615
            Left            =   240
            TabIndex        =   61
            Top             =   960
            Width           =   5535
         End
         Begin VB.Label Label42 
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
            Left            =   2160
            TabIndex        =   94
            Top             =   480
            Width           =   120
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Whatsapp Number"
            Height          =   375
            Left            =   240
            TabIndex        =   69
            Top             =   1560
            Width           =   2280
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Phone Number"
            Height          =   375
            Left            =   240
            TabIndex        =   68
            Top             =   600
            Width           =   1860
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Email Id"
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   2640
            Width           =   1050
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Emergency Number"
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   2040
            Width           =   2430
         End
      End
      Begin VB.Frame frmAddr 
         Caption         =   "Address"
         Height          =   4095
         Left            =   5640
         TabIndex        =   42
         Top             =   240
         Width           =   7335
         Begin VB.CheckBox Check3 
            Caption         =   "Tick if Correspondence Address is not same as Permanent Address"
            Height          =   615
            Left            =   240
            TabIndex        =   51
            Top             =   3360
            Width           =   6615
         End
         Begin VB.TextBox Text4 
            Height          =   495
            Left            =   5640
            TabIndex        =   50
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox Text16 
            Height          =   495
            Left            =   240
            TabIndex        =   49
            Top             =   2760
            Width           =   1815
         End
         Begin VB.TextBox Text15 
            Height          =   495
            Left            =   2160
            TabIndex        =   48
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text14 
            Height          =   495
            Left            =   3960
            TabIndex        =   47
            Top             =   2760
            Width           =   1575
         End
         Begin VB.TextBox Text13 
            Height          =   495
            Left            =   1800
            TabIndex        =   46
            Top             =   1800
            Width           =   5295
         End
         Begin VB.TextBox Text12 
            Height          =   495
            Left            =   1800
            TabIndex        =   45
            Top             =   1320
            Width           =   5295
         End
         Begin VB.TextBox Text11 
            Height          =   495
            Left            =   1800
            TabIndex        =   44
            Top             =   840
            Width           =   5295
         End
         Begin VB.TextBox Text10 
            Height          =   495
            Left            =   1800
            TabIndex        =   43
            Top             =   360
            Width           =   5295
         End
         Begin VB.Label Label24 
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
            Left            =   6840
            TabIndex        =   92
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label23 
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
            Left            =   5280
            TabIndex        =   91
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label17 
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
            Left            =   3360
            TabIndex        =   90
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label16 
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
            Left            =   1680
            TabIndex        =   89
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label15 
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
            Left            =   1440
            TabIndex        =   88
            Top             =   1800
            Width           =   120
         End
         Begin VB.Label Label7 
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
            Left            =   1080
            TabIndex        =   87
            Top             =   360
            Width           =   120
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Line 1"
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   480
            Width           =   765
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "District"
            Height          =   375
            Left            =   720
            TabIndex        =   58
            Top             =   2400
            Width           =   900
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "State"
            Height          =   375
            Left            =   2760
            TabIndex        =   57
            Top             =   2400
            Width           =   690
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Country"
            Height          =   375
            Left            =   4200
            TabIndex        =   56
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Pincode"
            Height          =   375
            Left            =   5880
            TabIndex        =   55
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Line 2"
            Height          =   375
            Left            =   240
            TabIndex        =   54
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Line 3"
            Height          =   375
            Left            =   240
            TabIndex        =   53
            Top             =   1440
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Landmark"
            Height          =   375
            Left            =   240
            TabIndex        =   52
            Top             =   1920
            Width           =   1200
         End
      End
      Begin VB.Frame fmCorspAddr 
         Caption         =   "Correspondence Address"
         Height          =   3975
         Left            =   5640
         TabIndex        =   23
         Top             =   240
         Width           =   7335
         Begin VB.CheckBox Check4 
            Caption         =   "View Address"
            Height          =   615
            Left            =   4560
            TabIndex        =   33
            Top             =   3240
            Width           =   4095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Same As Permanent Address"
            Height          =   615
            Left            =   240
            TabIndex        =   32
            Top             =   3240
            Width           =   3855
         End
         Begin VB.TextBox Text28 
            Height          =   495
            Left            =   1800
            TabIndex        =   31
            Top             =   360
            Width           =   5295
         End
         Begin VB.TextBox Text27 
            Height          =   495
            Left            =   1800
            TabIndex        =   30
            Top             =   840
            Width           =   5295
         End
         Begin VB.TextBox Text26 
            Height          =   495
            Left            =   1800
            TabIndex        =   29
            Top             =   1320
            Width           =   5295
         End
         Begin VB.TextBox Text25 
            Height          =   495
            Left            =   1800
            TabIndex        =   28
            Top             =   1800
            Width           =   5295
         End
         Begin VB.TextBox Text24 
            Height          =   495
            Left            =   3960
            TabIndex        =   27
            Top             =   2760
            Width           =   1575
         End
         Begin VB.TextBox Text23 
            Height          =   495
            Left            =   2160
            TabIndex        =   26
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox Text22 
            Height          =   495
            Left            =   240
            TabIndex        =   25
            Top             =   2760
            Width           =   1815
         End
         Begin VB.TextBox Text21 
            Height          =   495
            Left            =   5640
            TabIndex        =   24
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label50 
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
            Left            =   6840
            TabIndex        =   104
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label49 
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
            Left            =   5280
            TabIndex        =   103
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label48 
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
            Left            =   3360
            TabIndex        =   102
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label47 
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
            Left            =   1440
            TabIndex        =   101
            Top             =   1800
            Width           =   120
         End
         Begin VB.Label Label46 
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
            Left            =   1080
            TabIndex        =   100
            Top             =   360
            Width           =   120
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Landmark"
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   1920
            Width           =   1200
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Line 3"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   1440
            Width           =   765
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Line 2"
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Pincode"
            Height          =   375
            Left            =   5880
            TabIndex        =   38
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Country"
            Height          =   375
            Left            =   4200
            TabIndex        =   37
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "State"
            Height          =   375
            Left            =   2760
            TabIndex        =   36
            Top             =   2400
            Width           =   690
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "District"
            Height          =   375
            Left            =   720
            TabIndex        =   35
            Top             =   2400
            Width           =   900
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Line 1"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   480
            Width           =   765
         End
      End
      Begin VB.Frame frmGndr 
         Caption         =   "Gender"
         Height          =   975
         Left            =   360
         TabIndex        =   18
         Top             =   3600
         Width           =   5055
         Begin VB.OptionButton Option4 
            Caption         =   "Others"
            Height          =   495
            Left            =   3720
            TabIndex        =   22
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Trans"
            Height          =   495
            Left            =   2520
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Female"
            Height          =   495
            Left            =   1200
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Male"
            Height          =   495
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   3000
         TabIndex        =   16
         Top             =   3120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Format          =   125894657
         CurrentDate     =   45635
      End
      Begin VB.ComboBox Combo5 
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Text            =   " Department Id"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox Combo8 
         Height          =   495
         Left            =   3000
         TabIndex        =   2
         Text            =   " Position Id"
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   3000
         TabIndex        =   1
         Top             =   2520
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   13440
         TabIndex        =   75
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Format          =   125894657
         CurrentDate     =   45635
      End
      Begin VB.Label Label45 
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
         Left            =   8040
         TabIndex        =   99
         Top             =   4440
         Width           =   120
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "PAN"
         Height          =   375
         Left            =   9120
         TabIndex        =   98
         Top             =   4560
         Width           =   570
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Aadhar"
         Height          =   375
         Left            =   7080
         TabIndex        =   97
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label25 
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
         Left            =   15360
         TabIndex        =   93
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Employee Status"
         Height          =   375
         Left            =   13440
         TabIndex        =   72
         Top             =   3240
         Width           =   1995
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Employee Type"
         Height          =   375
         Left            =   13440
         TabIndex        =   71
         Top             =   1920
         Width           =   1905
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Joining"
         Height          =   375
         Left            =   13440
         TabIndex        =   70
         Top             =   480
         Width           =   1890
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   3120
         Width           =   1560
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
         Left            =   1800
         TabIndex        =   14
         Top             =   1920
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
         Left            =   1800
         TabIndex        =   13
         Top             =   840
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
         Left            =   2040
         TabIndex        =   12
         Top             =   1320
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
         TabIndex        =   11
         Top             =   360
         Width           =   120
      End
      Begin VB.Label lblDEPT_ID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3000
         TabIndex        =   10
         Top             =   1455
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "First Name"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Employee Id"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department Id"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Position Id"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Last Name"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2520
         Width           =   1245
      End
   End
   Begin MSDataGridLib.DataGrid DataGridEmp 
      Height          =   3855
      Left            =   120
      TabIndex        =   134
      Top             =   4800
      Visible         =   0   'False
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   6800
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
   Begin VB.Line Line4 
      BorderColor     =   &H00400040&
      BorderWidth     =   4
      X1              =   9120
      X2              =   11160
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00400040&
      BorderWidth     =   4
      X1              =   4200
      X2              =   6360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "EMPLOYEE"
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
      Left            =   6480
      TabIndex        =   15
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim toogle As Boolean



Private Sub cmdEmpSearch_Click()
toogle = False
frmEmpEntry.Visible = False
frmCtrlPnl.Visible = False
frmNavPnl.Visible = False
DataGridEmp.Visible = True
cmdEmpDisp.Caption = "&Back"
frmDispPnl.Top = 2520
frmSearch.Visible = True
End Sub

Private Sub Form_Load()
toogle = True
End Sub
Private Sub cmdEmpDisp_Click()
toogle = Not toogle
If toogle = False Then
frmEmpEntry.Visible = False
frmCtrlPnl.Visible = False
frmNavPnl.Visible = False
DataGridEmp.Visible = True
cmdEmpDisp.Caption = "&Back"
frmDispPnl.Top = 2520
frmSearch.Visible = True

Else

frmEmpEntry.Visible = True
frmCtrlPnl.Visible = True
frmNavPnl.Visible = True
DataGridEmp.Visible = False
cmdEmpDisp.Caption = "&Display Records"
frmDispPnl.Top = 4920
frmSearch.Visible = False

End If

End Sub

