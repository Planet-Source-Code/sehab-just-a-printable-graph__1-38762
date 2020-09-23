VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make A Line Graph"
   ClientHeight    =   5415
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VScroll1 
      Height          =   5175
      LargeChange     =   100
      Left            =   5160
      Max             =   4290
      SmallChange     =   100
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   0
      Max             =   4170
      SmallChange     =   100
      TabIndex        =   1
      Top             =   5160
      Width           =   5175
   End
   Begin VB.PictureBox Graph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H00000000&
      Height          =   5175
      Left            =   0
      ScaleHeight     =   345
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Text            =   "No Name"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   960
         Top             =   1080
      End
      Begin MSComDlg.CommonDialog Gra 
         Left            =   1560
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.gra"
         FontBold        =   -1  'True
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1320
         Top             =   2040
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   2160
         Top             =   1200
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnumang 
         Caption         =   "Make A New Graph"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuspace 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save Graph"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuog 
         Caption         =   "Open Graph"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuspacetwo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print Preview"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnumode 
         Caption         =   "Modes"
         Begin VB.Menu mnuline 
            Caption         =   "Line Graph"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnubar 
            Caption         =   "Bar Graph"
         End
      End
      Begin VB.Menu mnuspac 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim A As Integer
Dim C As Integer
Dim D As Integer
Dim E As Integer
Dim Dot(0 To 9) As Integer
Dim LineGraph As Boolean

Private Sub Form_Load()
'Call Draw_Graph
LineGraph = True
'D = 331
Graph.Picture = LoadPicture(App.Path & "\original.gra")
Graph.Width = 10000
Graph.Height = 10000
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HScroll1_Change()
Graph.Left = "-" & HScroll1.Value
End Sub

Private Sub mnubar_Click()
LineGraph = False
mnubar.Checked = True
mnuline.Checked = False
Me.Caption = "Make A Bar Graph"
Graph.Cls
Timer3.Enabled = True
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuline_Click()
LineGraph = True
mnuline.Checked = True
mnubar.Checked = False
Me.Caption = "Make A Line Graph"
Graph.Cls
Timer3.Enabled = True
End Sub

Private Sub mnumang_Click()
Form2.Show
End Sub

Private Sub mnuog_Click()
Dim saved As String
Gra.ShowOpen
Dot(0) = 1000 - Form2.Text1(0).Text
Dot(0) = Dot(0) / 3.3
Graph.Picture = LoadPicture(Gra.FileName)
Gra.Filter = "*.bmp"
End Sub

Private Sub mnuprint_Click()
Graph.Visible = False
Form3.Show
Form3.CurrentX = 5
Form3.CurrentY = 35
SavePicture Graph.Image, App.Path & "\cc.gra"
Form3.Picture = LoadPicture(App.Path & "\cc.gra")
Form3.CurrentX = 150
Form3.CurrentY = 360
Form3.Print Text1.Text
Graph.Visible = True
End Sub

'Private Sub Draw_Graph()
'C = C + 30
'Graph.ForeColor = &HE0E0E0
'Graph.Line (C, 0)-(C, 300)
'If C = 11 * 30 Then Timer1.Enabled = False
'End Sub

Private Sub mnusave_Click()
Dim saved As String
Gra.ShowSave
saved = Gra.FileName
Gra.Filter = "*.bmp"
SavePicture Graph.Image, saved
End Sub

'Private Sub Timer1_Timer()
'Call Draw_Graph
'End Sub

'Private Sub Timer2_Timer()
'D = D - 30
'If D = 300 / 30 Then Timer2.Enabled = False
'Graph.Line (0, D)-(330, D)
'E = E + 100
'Graph.CurrentX = 0
'Graph.CurrentY = D - 5
'Graph.ForeColor = vbBlack
'Graph.Print E
'Graph.ForeColor = &HE0E0E0
'End Sub

Private Sub Timer3_Timer()
If LineGraph = True Then
Graph.ForeColor = vbBlack
Dot(0) = 1000 - Form2.Text1(0).Text
Dot(0) = Dot(0) / 3.3
Graph.Circle (45, Dot(0)), 3, 21
Graph.CurrentX = 50
Graph.CurrentY = Dot(0)
Graph.Print Form2.Text1(0).Text

Dot(1) = 1000 - Form2.Text1(1).Text
Dot(1) = Dot(1) / 3.3
Graph.Circle (75, Dot(1)), 3, 21
Graph.CurrentX = 80
Graph.CurrentY = Dot(1)
Graph.Print Form2.Text1(1).Text

Dot(2) = 1000 - Form2.Text1(2).Text
Dot(2) = Dot(2) / 3.3
Graph.Circle (105, Dot(2)), 3, 21
Graph.CurrentX = 110
Graph.CurrentY = Dot(2)
Graph.Print Form2.Text1(2).Text

Dot(3) = 1000 - Form2.Text1(3).Text
Dot(3) = Dot(3) / 3.3
Graph.Circle (135, Dot(3)), 3, 21
Graph.CurrentX = 140
Graph.CurrentY = Dot(3)
Graph.Print Form2.Text1(3).Text

Dot(4) = 1000 - Form2.Text1(4).Text
Dot(4) = Dot(4) / 3.3
Graph.Circle (165, Dot(4)), 3, 21
Graph.CurrentX = 170
Graph.CurrentY = Dot(4)
Graph.Print Form2.Text1(4).Text

Dot(5) = 1000 - Form2.Text1(5).Text
Dot(5) = Dot(5) / 3.3
Graph.Circle (195, Dot(5)), 3, 21
Graph.CurrentX = 200
Graph.CurrentY = Dot(5)
Graph.Print Form2.Text1(5).Text

Dot(6) = 1000 - Form2.Text1(6).Text
Dot(6) = Dot(6) / 3.3
Graph.Circle (225, Dot(6)), 3, 21
Graph.CurrentX = 230
Graph.CurrentY = Dot(6)
Graph.Print Form2.Text1(6).Text

Dot(7) = 1000 - Form2.Text1(7).Text
Dot(7) = Dot(7) / 3.3
Graph.Circle (255, Dot(7)), 3, 21
Graph.CurrentX = 260
Graph.CurrentY = Dot(7)
Graph.Print Form2.Text1(7).Text

Dot(8) = 1000 - Form2.Text1(8).Text
Dot(8) = Dot(8) / 3.3
Graph.Circle (285, Dot(8)), 3, 21
Graph.CurrentX = 290
Graph.CurrentY = Dot(8)
Graph.Print Form2.Text1(8).Text

Dot(9) = 1000 - Form2.Text1(9).Text
Dot(9) = Dot(9) / 3.3
Graph.Circle (315, Dot(9)), 3, 21
Graph.CurrentX = 320
Graph.CurrentY = Dot(9)
Graph.Print Form2.Text1(9).Text

Graph.Line (45, Dot(0))-(75, Dot(1))
Graph.Line (75, Dot(1))-(105, Dot(2))
Graph.Line (105, Dot(2))-(135, Dot(3))
Graph.Line (135, Dot(3))-(165, Dot(4))
Graph.Line (165, Dot(4))-(195, Dot(5))
Graph.Line (195, Dot(5))-(225, Dot(6))
Graph.Line (225, Dot(6))-(255, Dot(7))
Graph.Line (255, Dot(7))-(285, Dot(8))
Graph.Line (285, Dot(8))-(315, Dot(9))
Timer3.Enabled = False
End If

If LineGraph = False Then
Dot(0) = 1000 - Form2.Text1(0).Text
Dot(0) = Dot(0) / 3.3
Dot(1) = 1000 - Form2.Text1(1).Text
Dot(1) = Dot(1) / 3.3
Dot(2) = 1000 - Form2.Text1(2).Text
Dot(2) = Dot(2) / 3.3
Dot(3) = 1000 - Form2.Text1(3).Text
Dot(3) = Dot(3) / 3.3
Dot(4) = 1000 - Form2.Text1(4).Text
Dot(4) = Dot(4) / 3.3
Dot(5) = 1000 - Form2.Text1(5).Text
Dot(5) = Dot(5) / 3.3
Dot(6) = 1000 - Form2.Text1(6).Text
Dot(6) = Dot(6) / 3.3
Dot(7) = 1000 - Form2.Text1(7).Text
Dot(7) = Dot(7) / 3.3
Dot(8) = 1000 - Form2.Text1(8).Text
Dot(8) = Dot(8) / 3.3
Dot(9) = 1000 - Form2.Text1(9).Text
Dot(9) = Dot(9) / 3.3

Graph.ForeColor = Form2.Text1(0).BackColor
Graph.Line (30, 300)-(60, Dot(0)), , BF
Graph.CurrentX = 35
Graph.CurrentY = Dot(0) + 5
Graph.Print Form2.Text1(0).Text

Graph.ForeColor = Form2.Text1(1).BackColor
Graph.Line (60, 300)-(90, Dot(1)), , BF
Graph.CurrentX = 65
Graph.CurrentY = Dot(1) + 5
Graph.Print Form2.Text1(1).Text

Graph.ForeColor = Form2.Text1(2).BackColor
Graph.Line (90, 300)-(120, Dot(2)), , BF
Graph.CurrentX = 100
Graph.CurrentY = Dot(2) + 5
Graph.Print Form2.Text1(2).Text

Graph.ForeColor = Form2.Text1(3).BackColor
Graph.Line (120, 300)-(150, Dot(3)), , BF
Graph.CurrentX = 125
Graph.CurrentY = Dot(3) + 5
Graph.Print Form2.Text1(3).Text

Graph.ForeColor = Form2.Text1(4).BackColor
Graph.Line (150, 300)-(180, Dot(4)), , BF
Graph.CurrentX = 155
Graph.CurrentY = Dot(4) + 5
Graph.Print Form2.Text1(4).Text

Graph.ForeColor = Form2.Text1(5).BackColor
Graph.Line (180, 300)-(210, Dot(5)), , BF
Graph.CurrentX = 185
Graph.CurrentY = Dot(5) + 5
Graph.Print Form2.Text1(5).Text

Graph.ForeColor = Form2.Text1(6).BackColor
Graph.Line (210, 300)-(240, Dot(6)), , BF
Graph.CurrentX = 215
Graph.CurrentY = Dot(6) + 5
Graph.Print Form2.Text1(6).Text

Graph.ForeColor = Form2.Text1(7).BackColor
Graph.Line (240, 300)-(270, Dot(7)), , BF
Graph.CurrentX = 245
Graph.CurrentY = Dot(7) + 5
Graph.Print Form2.Text1(7).Text

Graph.ForeColor = Form2.Text1(8).BackColor
Graph.Line (270, 300)-(300, Dot(8)), , BF
Graph.CurrentX = 275
Graph.CurrentY = Dot(8) + 5
Graph.Print Form2.Text1(8).Text

Graph.ForeColor = Form2.Text1(9).BackColor
Graph.Line (300, 300)-(330, Dot(9)), , BF
Graph.CurrentX = 305
Graph.CurrentY = Dot(9) + 5
Graph.Print Form2.Text1(9).Text
Timer3.Enabled = False

End If

Graph.ForeColor = vbBlack
If LineGraph = False Then
Graph.CurrentX = 150
Graph.CurrentY = 350
Graph.Print "Bar Graph"
Else
Graph.CurrentX = 150
Graph.CurrentY = 350
Graph.Print "Line Graph"
End If

Graph.CurrentX = 30
Graph.CurrentY = 320
Graph.Print Form2.Text2.Text

End Sub

Private Sub VScroll1_Change()
Graph.Top = "-" & VScroll1.Value
End Sub
