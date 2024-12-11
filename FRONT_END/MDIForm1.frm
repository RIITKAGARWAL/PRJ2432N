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
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   15135
      TabIndex        =   1
      Top             =   0
      Width           =   15195
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Height          =   7755
      Left            =   0
      ScaleHeight     =   7695
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   735
      Width           =   4400
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
'Form1.Show
'Form4.Show
'Form5.Show
'Form8.Show
'Form9.Show
'Form5.Show
End Sub

Private Sub mnuBankDetails_Click()
frmBank.Show
End Sub

Private Sub mnuDepartment_Click()
frmDepartment.Show
End Sub

Private Sub mnuEmployee_Click()
frmEmp.Show
End Sub

Private Sub mnuPosition_Click()
frmPosition.Show
End Sub
