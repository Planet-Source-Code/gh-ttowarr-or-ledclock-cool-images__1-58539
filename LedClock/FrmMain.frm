VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Clock"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   8
      Left            =   -480
      ScaleHeight     =   3105
      ScaleWidth      =   585
      TabIndex        =   14
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   11025
      TabIndex        =   11
      Top             =   -120
      Width           =   11055
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   7
      Left            =   3480
      ScaleHeight     =   3135
      ScaleWidth      =   15
      TabIndex        =   13
      Top             =   -120
      Width           =   15
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   1
      Left            =   3480
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   3135
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   3
      Left            =   7200
      ScaleHeight     =   3105
      ScaleWidth      =   0
      TabIndex        =   9
      Top             =   -120
      Width           =   15
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   6
      Left            =   7200
      Picture         =   "FrmMain.frx":93EA
      ScaleHeight     =   3135
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   4
      Left            =   9120
      ScaleHeight     =   3105
      ScaleWidth      =   345
      TabIndex        =   10
      Top             =   -240
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   2
      Left            =   5400
      ScaleHeight     =   3105
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   0
      Left            =   1680
      ScaleHeight     =   3105
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Tmr 
      Interval        =   1
      Left            =   10440
      Top             =   2280
   End
   Begin PrjClock.Led Led 
      Height          =   2835
      Index           =   3
      Left            =   5520
      TabIndex        =   3
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   5001
   End
   Begin PrjClock.Led Led 
      Height          =   2835
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   5001
   End
   Begin PrjClock.Led Led 
      Height          =   2835
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   5001
   End
   Begin PrjClock.Led Led 
      Height          =   2835
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   5001
   End
   Begin PrjClock.Led Led 
      Height          =   2835
      Index           =   4
      Left            =   9240
      TabIndex        =   4
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   5001
   End
   Begin PrjClock.Led Led 
      Height          =   2835
      Index           =   5
      Left            =   7440
      TabIndex        =   5
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   5001
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Tmr_Timer()
If Len(Hour(Now)) = 2 Then
Led(0).SetNumber Mid$(Hour(Now), 1, 1)
Led(1).SetNumber Mid$(Hour(Now), 2, 1)
Else
Led(1).SetNumber Mid$(Hour(Now), 1, 1)
End If
If Len(Minute(Now)) = 2 Then
Led(2).SetNumber Mid$(Minute(Now), 1, 1)
Led(3).SetNumber Mid$(Minute(Now), 2, 1)
Else
Led(3).SetNumber Mid$(Minute(Now), 1, 1)
End If
If Len(Second(Now)) = 2 Then
Led(4).SetNumber Mid$(Second(Now), 2, 1)
Led(5).SetNumber Mid$(Second(Now), 1, 1)
Else
Led(4).SetNumber Mid$(Second(Now), 1, 1)
End If
End Sub
