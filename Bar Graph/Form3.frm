VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preview And Print"
   ClientHeight    =   4830
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4710
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnupr 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnupr_Click()
PrintForm
End Sub

