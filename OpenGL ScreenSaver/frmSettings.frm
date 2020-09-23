VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3135
   ClientLeft      =   3540
   ClientTop       =   2835
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   840
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtDepth 
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Text            =   "0.2"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Text"
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3495
         Begin VB.CommandButton cmdFont 
            Caption         =   "Font..."
            Height          =   375
            Left            =   960
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox cboTimeFormat 
            Height          =   315
            Left            =   1680
            TabIndex        =   10
            Text            =   "Combo1"
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton optDisp 
            Caption         =   "Time"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optDisp 
            Caption         =   "Text"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txt3D 
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Text            =   "OpenGL Tim"
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame frame2 
         Caption         =   "Rotate Around"
         Height          =   735
         Left            =   3720
         TabIndex        =   4
         Top             =   120
         Width           =   1815
         Begin VB.ComboBox cboRotate 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label Label1 
         Caption         =   "3D Depth"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
  RotAxis = cboRotate.ListIndex + 1
  sTime = optDisp(1).Value
  TimeFormat = cboTimeFormat
  sText = txt3D
  sFont = txt3D.FontName
  sDepth = txtDepth

  SaveConfig
  cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFont_Click()
  With cDialog
    .flags = cdlCFScreenFonts
    .FontName = txt3D.Font
    .ShowFont
    txt3D.Font = .FontName
    cmdFont.ToolTipText = .FontName
  End With
End Sub

Private Sub cmdOk_Click()
  cmdApply_Click
  cmdCancel_Click
End Sub

Private Sub Form_Load()
  LoadConfig
  
  txt3D = sText
  txt3D.FontName = sFont
  txtDepth = sDepth
  cmdFont.ToolTipText = sFont
  
  optDisp(0).Value = Not sTime
  optDisp(1).Value = sTime
  
  With cboTimeFormat
    .AddItem "HH:MM:SS"
    .AddItem "HH-MM-SS"
    .AddItem "HH MM SS"
    .AddItem "HHMMSS"
    .AddItem "HH:MM"
    .AddItem "HH-MM"
    .AddItem "HH MM"
    .AddItem "HHMM"
    .Text = TimeFormat
  End With
  
  With cboRotate
    .AddItem "X Axis"
    .AddItem "Y Axis"
    .AddItem "Z Axis"
    .AddItem "All"
    .ListIndex = RotAxis - 1
  End With
End Sub

Private Sub optDisp_Click(Index As Integer)
  If Index = 0 Then
    cboTimeFormat.Enabled = False
    txt3D.Enabled = True
  Else
    cboTimeFormat.Enabled = True
    txt3D.Enabled = False
  End If
End Sub

Private Sub txtDepth_Change()
  If Val(txtDepth) < 0 Then txtDepth = 0
  If Val(txtDepth) > 50 Then txtDepth = 50
End Sub
