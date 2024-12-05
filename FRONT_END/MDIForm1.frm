VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "AIC SOLUTIONS - A PERSONALIZED SOFTWARE FOR AGARWAL'S INVEST CARE"
   ClientHeight    =   8490
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15195
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      Picture         =   "MDIForm1.frx":2F8DE
      ScaleHeight     =   555
      ScaleWidth      =   15135
      TabIndex        =   0
      Top             =   0
      Width           =   15195
   End
   Begin VB.Menu Department 
      Caption         =   "&Department"
      Begin VB.Menu mnuDepartment 
         Caption         =   "&Department"
      End
      Begin VB.Menu mnuPosition 
         Caption         =   "&Position"
      End
   End
   Begin VB.Menu Employee 
      Caption         =   "&Employee"
      Begin VB.Menu mnuEmployee 
         Caption         =   "&Employee"
      End
      Begin VB.Menu mnuBankDetails 
         Caption         =   "&Bank Details"
      End
   End
   Begin VB.Menu Attendance 
      Caption         =   "&Attendance And Leave"
      Begin VB.Menu mnuLeave 
         Caption         =   "&Leave"
      End
      Begin VB.Menu mnuAttendance 
         Caption         =   "&Attendance"
      End
   End
   Begin VB.Menu mnuPayroll 
      Caption         =   "&Payroll"
   End
   Begin VB.Menu mnuBonus 
      Caption         =   "&Bonus And Incentives"
   End
   Begin VB.Menu mnuAdvance 
      Caption         =   "Advance (&Loan)"
   End
   Begin VB.Menu mnuRetirement 
      Caption         =   "&Retirement And Benefits"
   End
   Begin VB.Menu mnuTaxation 
      Caption         =   "&Taxation And Compliance"
   End
   Begin VB.Menu Salary 
      Caption         =   "&Salary"
      Begin VB.Menu mnuPaySlip 
         Caption         =   "&Pay Slip Generation"
      End
      Begin VB.Menu mnuSalary 
         Caption         =   "&Salary"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Form1.Show
Form3.Show
End Sub
