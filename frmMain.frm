VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Loading - please wait"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2880
      Top             =   1080
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Timer1_Timer()

' Using a standard timer control for traffic light changes
booChange = True

End Sub



