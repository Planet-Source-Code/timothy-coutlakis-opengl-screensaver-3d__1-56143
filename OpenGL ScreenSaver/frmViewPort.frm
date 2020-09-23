VERSION 5.00
Begin VB.Form frmViewPort 
   BackColor       =   &H00FBAD44&
   BorderStyle     =   0  'None
   ClientHeight    =   2910
   ClientLeft      =   2010
   ClientTop       =   2430
   ClientWidth     =   3480
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   194
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   232
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label lblOGL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OpenGL Text"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1545
   End
End
Attribute VB_Name = "frmViewPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
  If RunMode = rmScreenSaver Then Unload Me
End Sub

Private Sub Form_DblClick()
  If RunMode = rmScreenSaver Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If RunMode = rmScreenSaver Then Unload Me
  Keys(KeyCode) = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If RunMode = rmScreenSaver Then Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Keys(KeyCode) = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If RunMode = rmScreenSaver Then Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Do nothing except in screen saver mode.
  If RunMode <> rmScreenSaver Then Exit Sub

  Static Xlast, Ylast
  Dim Xnow As Single
  Dim Ynow As Single
  Xnow = X
  Ynow = Y
  If Xlast = 0 And Ylast = 0 Then
    Xlast = Xnow
    Ylast = Ynow
  End If
  If (Xnow <> Xlast Or Ynow <> Ylast) Then Unload Me
End Sub


Private Sub Form_Resize()
  ' Load configuration information.
  LoadConfig
  'ReSizeGLScene ScaleWidth, ScaleHeight
  
  lblOGL.Left = ScaleWidth / 2 - lblOGL.Width / 2
  lblOGL.Top = ScaleHeight / 2 - lblOGL.height / 2
End Sub

' Redisplay the cursor if we hid it in Sub Main.
Private Sub Form_Unload(Cancel As Integer)
  If RunMode = rmScreenSaver Then ShowCursor True
  KillGLWindow
End Sub

' Note the ScaleMode of this form is set to pixels
