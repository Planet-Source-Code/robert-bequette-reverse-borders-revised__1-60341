VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Reverse Borders"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Standard Dialog"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Standard TaskBar"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reverse TaskBar"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reverse Tool Window"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tool Window"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reverse Dialog"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In order to see the effects you must restart the project after each effect is applied to see a diff one no biggy.
Private Sub Command1_Click()
MTBackWardsWindow Me.hwnd
End Sub

Private Sub Command2_Click()
MTToolWindow Me.hwnd
End Sub

Private Sub Command3_Click()
MTBackWardsToolWindow Me.hwnd
End Sub

Private Sub Command4_Click()
Dim shelltraywnd As Long
shelltraywnd = FindWindow("Shell_TrayWnd", vbNullString)
MTBackWardsWindow shelltraywnd
End Sub

Private Sub Command5_Click()
Dim shelltraywnd As Long
shelltraywnd = FindWindow("Shell_TrayWnd", vbNullString)
MTStandardWindow shelltraywnd
End Sub

Private Sub Command6_Click()
MTStandardWindow Me.hwnd
End Sub

Private Sub Form_Load()
SOT Me
End Sub
