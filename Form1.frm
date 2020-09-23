VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Waverizer by MTECH Designs."
   ClientHeight    =   7335
   ClientLeft      =   1890
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   Begin VB.HScrollBar HScroll3 
      Height          =   285
      Left            =   135
      Max             =   100
      Min             =   -100
      TabIndex        =   6
      Top             =   6975
      Value           =   10
      Width           =   4815
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   285
      Left            =   135
      Max             =   15
      Min             =   -15
      TabIndex        =   4
      Top             =   6660
      Value           =   1
      Width           =   4830
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   135
      Max             =   1000
      TabIndex        =   3
      Top             =   6315
      Width           =   6285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Main Loop!!!! Only 7 Lines !!!!!!"
      Height          =   360
      Left            =   585
      TabIndex        =   2
      Top             =   5445
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.PictureBox picSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   6510
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   419
      TabIndex        =   1
      Top             =   1230
      Visible         =   0   'False
      Width           =   6285
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5955
      Left            =   30
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   419
      TabIndex        =   0
      Top             =   -45
      Width           =   6345
   End
   Begin VB.Label Label3 
      Caption         =   "Use this scroller to make the wave ripple..."
      Height          =   285
      Left            =   135
      TabIndex        =   8
      Top             =   6060
      Width           =   6210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amplitude: 10"
      Height          =   195
      Left            =   4995
      TabIndex        =   7
      Top             =   7020
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Waves: 1"
      Height          =   195
      Left            =   5010
      TabIndex        =   5
      Top             =   6705
      Width           =   1470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public X As Long, Y As Long, Ang As Long, Dep As Long, Mult As Long, Pi As Variant, Wave As Variant, Heig As Long, Wid As Long
'WAVERIZER by MTECH Designs (Michael Pote)
'-----------------------------------------
'Only 7 Lines of code to make a big picture occilate!
'Mult - Number of waves
'Ang - Angle of Sine wave
'Dep - Amplitude / Depth

Private Sub Command1_Click()
For X = 1 To Wid
DoEvents
Ang = HScroll1.Value + (X * Mult)
Wave = Sin(Pi * Ang) * Dep
BitBlt picDest.hdc, X, Wave, 1, Heig, picSrc.hdc, X, 0, SRCCOPY
Next
picDest.Refresh
End Sub

Private Sub Form_Load()
'To make the loop as fast as possible, I put
'all the preset things that will never change
'in form_load
Dep = 30
Pi = (3.1456 / 180)
Heig = picDest.ScaleHeight
Wid = picDest.ScaleWidth
Mult = 1
End Sub

Private Sub HScroll1_Scroll()
Command1_Click

End Sub

Private Sub HScroll2_Change()
Mult = HScroll2.Value
Command1_Click
Label1.Caption = "Number of Waves: " & Mult
End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change
End Sub

Private Sub HScroll3_Change()
HScroll3_Scroll
End Sub

Private Sub HScroll3_Scroll()
Dep = HScroll3.Value
Label2.Caption = "Amplitude: " & HScroll3.Value
Command1_Click
End Sub
