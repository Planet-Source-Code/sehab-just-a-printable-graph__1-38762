VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Make New Graph"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form2"
   ScaleHeight     =   2850
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   14
      Text            =   "Name Of The Graph"
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Graph"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      Height          =   285
      Index           =   9
      Left            =   4440
      TabIndex        =   10
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Index           =   8
      Left            =   3960
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      Height          =   285
      Index           =   7
      Left            =   3480
      TabIndex        =   8
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000000FF&
      Height          =   285
      Index           =   6
      Left            =   3000
      TabIndex        =   7
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   6
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C000C0&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.HScrollBar B 
      Height          =   255
      Left            =   120
      Max             =   9
      TabIndex        =   0
      Top             =   840
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   4635
      TabIndex        =   13
      Top             =   1680
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim A As Integer

Private Sub B_Change()
If B = 0 Then
Text1(0).Visible = True
Text1(1).Visible = False
Text1(2).Visible = False
Text1(3).Visible = False
Text1(4).Visible = False
Text1(5).Visible = False
Text1(6).Visible = False
Text1(7).Visible = False
Text1(8).Visible = False
Text1(9).Visible = False
End If

If B = 1 Then
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = False
Text1(3).Visible = False
Text1(4).Visible = False
Text1(5).Visible = False
Text1(6).Visible = False
Text1(7).Visible = False
Text1(8).Visible = False
Text1(9).Visible = False
End If

If B = 2 Then
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = True
Text1(3).Visible = False
Text1(4).Visible = False
Text1(5).Visible = False
Text1(6).Visible = False
Text1(7).Visible = False
Text1(8).Visible = False
Text1(9).Visible = False
End If

If B = 3 Then
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = True
Text1(3).Visible = True
Text1(4).Visible = False
Text1(5).Visible = False
Text1(6).Visible = False
Text1(7).Visible = False
Text1(8).Visible = False
Text1(9).Visible = False
End If

If B = 4 Then
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = True
Text1(3).Visible = True
Text1(4).Visible = True
Text1(5).Visible = False
Text1(6).Visible = False
Text1(7).Visible = False
Text1(8).Visible = False
Text1(9).Visible = False
End If

If B = 5 Then
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = True
Text1(3).Visible = True
Text1(4).Visible = True
Text1(5).Visible = True
Text1(6).Visible = False
Text1(7).Visible = False
Text1(8).Visible = False
Text1(9).Visible = False
End If

If B = 6 Then
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = True
Text1(3).Visible = True
Text1(4).Visible = True
Text1(5).Visible = True
Text1(6).Visible = True
Text1(7).Visible = False
Text1(8).Visible = False
Text1(9).Visible = False
End If

If B = 7 Then
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = True
Text1(3).Visible = True
Text1(4).Visible = True
Text1(5).Visible = True
Text1(6).Visible = True
Text1(7).Visible = True
Text1(8).Visible = False
Text1(9).Visible = False
End If

If B = 8 Then
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = True
Text1(3).Visible = True
Text1(4).Visible = True
Text1(5).Visible = True
Text1(6).Visible = True
Text1(7).Visible = True
Text1(8).Visible = True
Text1(9).Visible = False
End If

If B = 9 Then
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = True
Text1(3).Visible = True
Text1(4).Visible = True
Text1(5).Visible = True
Text1(6).Visible = True
Text1(7).Visible = True
Text1(8).Visible = True
Text1(9).Visible = True
End If

End Sub

Private Sub Command1_Click()
Form1.Graph.Cls
Form1.Timer3.Enabled = True
If Text3.Text <> "Name Of The Graph" Or Text3.Text <> "" Then
Form1.Text1.Text = Text3.Text
End If
End Sub

Private Sub Text1_Change(Index As Integer)
For A = 0 To 9
If Text1(A).Text = "" Then
Text1(A).Text = ""
Else
If Text1(A).Text > 1000 Then
Text1(A).Text = 1000
End If
End If
Next A
End Sub
