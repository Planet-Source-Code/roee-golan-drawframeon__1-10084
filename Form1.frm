VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Draw Frame On"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   2790
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   3000
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1455
      Left            =   5040
      TabIndex        =   5
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DrawFrameOn(TopLeftControl As Control, LowestRightControl As Control, Style As String, Framewidth)
Dim dw, fs, sm
Dim st$
Dim Lft, Toplft, Hite
Dim Rite, Ritebotm
Dim lt As Long
Dim rb As Long
dw = DrawWidth
fs = FillStyle
sm = ScaleMode
DrawWidth = 1
FillStyle = 1
ScaleMode = 3
st = LCase(Left$(Style, 1))
Lft = TopLeftControl.Left
Toplft = TopLeftControl.Top
Hite = TopLeftControl.Height
Rite = LowestRightControl.Left + LowestRightControl.Width
Ritebotm = LowestRightControl.Top + LowestRightControl.Height
If Ritebotm > Hite Then Hite = Ritebotm
lt = vb3DHighlight
rb = vbButtonShadow
If st = "i" Then
lt = vb3DDKShadow
rb = vb3DHighlight
End If
Line (Lft - Framewidth, Toplft - Framewidth)-(Rite + Framewidth, Toplft - Framewidth), lt
Line (Lft - Framewidth, Toplft - Framewidth)-(Lft - Framewidth, Hite + Framewidth), lt
Line (Rite + Framewidth, Toplft - Framewidth)-(Rite + Framewidth, Ritebotm + Framewidth), rb
Line (Rite + Framewidth, Ritebotm + Framewidth)-(Lft - Framewidth, Hite + Framewidth), rb
DrawWidth = dw
FillStyle = fs
ScaleMode = sm
End Sub
Private Sub Form_Paint()
Text1.SetFocus
DrawFrameOn Text1, Text1, "outward", 5
DrawFrameOn Command1, Command1, "outward", 4
DrawFrameOn Command1, Command1, "inward", 1
DrawFrameOn Frame1, Frame1, "outward", 5
DrawFrameOn Option1, Option1, "outward", 5
DrawFrameOn Check1, Check1, "outward", 5
DrawFrameOn Check1, Check1, "inward", 2
DrawFrameOn VScroll1, VScroll1, "outward", 5
DrawFrameOn VScroll1, VScroll1, "inward", 2
DrawFrameOn Dir1, Dir1, "outward", 5
End Sub
