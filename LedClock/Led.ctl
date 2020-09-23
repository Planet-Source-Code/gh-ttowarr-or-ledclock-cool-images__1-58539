VERSION 5.00
Begin VB.UserControl Led 
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11475
   ScaleHeight     =   5910
   ScaleWidth      =   11475
   Begin VB.Image ImgLed 
      Height          =   2835
      Left            =   0
      Picture         =   "Led.ctx":0000
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   10
      Left            =   9600
      Picture         =   "Led.ctx":10A0A
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   9
      Left            =   7680
      Picture         =   "Led.ctx":21414
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   8
      Left            =   5760
      Picture         =   "Led.ctx":31E1E
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   7
      Left            =   3840
      Picture         =   "Led.ctx":42828
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   6
      Left            =   1920
      Picture         =   "Led.ctx":53232
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   5
      Left            =   0
      Picture         =   "Led.ctx":63C3C
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   4
      Left            =   9600
      Picture         =   "Led.ctx":74646
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   3
      Left            =   7680
      Picture         =   "Led.ctx":85050
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   2
      Left            =   5760
      Picture         =   "Led.ctx":95A5A
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   1
      Left            =   3840
      Picture         =   "Led.ctx":A6464
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image ImgNumber 
      Height          =   2835
      Index           =   0
      Left            =   1920
      Picture         =   "Led.ctx":B6E6E
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "Led"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetNumber(Number As Integer)
ImgLed.Picture = ImgNumber(Number).Picture
End Sub

Private Sub UserControl_Resize()
UserControl.Width = ImgLed.Width
UserControl.Height = ImgLed.Height
End Sub
