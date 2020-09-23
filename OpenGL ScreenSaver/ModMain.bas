Attribute VB_Name = "SubMain"
Option Explicit

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOP = 0

Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

' Global variables.
Public Const rmConfigure = 1
Public Const rmScreenSaver = 2
Public Const rmPreview = 3
Public RunMode As Integer

' Private variables.
Private Const APP_NAME = "Screen_Saver"

Public RotAxis As Integer
Public TimeFormat As String
Public sTime As Boolean
Public sText As String
Public sFont As String
Public sDepth As Single

' See if another instance of the program is
' running in screen saver mode.
Private Sub CheckShouldRun()
  ' If no instance is running, we're safe.
  If Not App.PrevInstance Then Exit Sub
  
  ' See if there is a screen saver mode instance.
  If FindWindow(vbNullString, APP_NAME) Then End
  
  ' Set our caption so other instances can find
  ' us in the previous line.
  frmViewPort.Caption = APP_NAME
End Sub

' Load configuration information from the registry.
Public Sub LoadConfig()
  RotAxis = GetSetting(App.Path, "Settings", "RotAxis", 4)
  TimeFormat = GetSetting(App.Path, "Settings", "TimeFormat", "HH:MM:SS")
  sTime = GetSetting(App.Path, "Settings", "sTime", False)
  sText = GetSetting(App.Path, "Settings", "sText", "OpenGL Tim")
  sFont = GetSetting(App.Path, "Settings", "sFont", "Comic Sans MS")
  sDepth = GetSetting(App.Path, "Settings", "sDepth", 0.2)
  'sFont = "courier new"
  'sText = "HHOHH" 'test text
  'sTime = False
  'RotAxis = 2
  'sDepth = 0.2
End Sub

' Save configuration information in the registry.
Public Sub SaveConfig()
  SaveSetting App.Path, "Settings", "RotAxis", RotAxis
  SaveSetting App.Path, "Settings", "TimeFormat", TimeFormat
  SaveSetting App.Path, "Settings", "sTime", sTime
  SaveSetting App.Path, "Settings", "sText", sText
  SaveSetting App.Path, "Settings", "sFont", sFont
  SaveSetting App.Path, "Settings", "sDepth", sDepth
End Sub

' Start the program.
Public Sub Main()

  'If You Want To Be Able To Configure The Executable,
  'Type (, "") after Case "/C" and
  'Delete (, "") after Case "/S"
  '
  'It Works!!
  
  Dim args As String
  Dim preview_hwnd As Long
  Dim preview_rect As RECT
  Dim window_style As Long

  ' Get the command line arguments.
  args = UCase$(Trim$(Command$))
  
  ' Examine the first 2 characters.
  Select Case Mid$(args, 1, 2)
    Case "/C" ', ""  ' Display configuration dialog.
      RunMode = rmConfigure
    Case "/S", ""    ' Run as a screen saver.
      RunMode = rmScreenSaver
    Case "/P"       ' Run in preview mode.
      RunMode = rmPreview
    Case Else       ' This shouldn't happen.
      RunMode = rmScreenSaver
  End Select

  Select Case RunMode
    Case rmConfigure    ' Display configuration dialog.
      frmSettings.Show vbModal
      
    Case rmScreenSaver  ' Run as a screen saver.
      ' Make sure there isn't another one running.
      CheckShouldRun

      Call GL_Main

    Case rmPreview      'Preview
      ' Get the preview area hWnd.
      preview_hwnd = GetHwndFromCommand(args)

      ' Get the dimensions of the preview area.
      GetClientRect preview_hwnd, preview_rect

      Load frmViewPort

      ' Set the caption for Windows 95.
      frmViewPort.Caption = "Preview"

      ' Get the current window style.
      window_style = GetWindowLong(frmViewPort.hwnd, GWL_STYLE)

      ' Add WS_CHILD to make this a child window.
      window_style = (window_style Or WS_CHILD)

      ' Set the window's new style.
      SetWindowLong frmViewPort.hwnd, _
          GWL_STYLE, window_style

      ' Set the window's parent so it appears
      ' inside the preview area.
      SetParent frmViewPort.hwnd, preview_hwnd

      ' Save the preview area's hWnd in
      ' the form's window structure.
      SetWindowLong frmViewPort.hwnd, _
          GWL_HWNDPARENT, preview_hwnd

      ' Show the preview.
      SetWindowPos frmViewPort.hwnd, _
          HWND_TOP, 0&, 0&, _
          preview_rect.Right, _
          preview_rect.Bottom, _
          SWP_NOZORDER Or SWP_NOACTIVATE Or _
              SWP_SHOWWINDOW
  End Select
End Sub

' Get the hWnd for the preview window from the
' command line arguments.
Private Function GetHwndFromCommand(ByVal args As String) As Long
  Dim argslen As Integer
  Dim i As Integer
  Dim ch As String

  ' get the numbers on the right of string
  args = Trim$(args)
  argslen = Len(args)
  For i = argslen To 1 Step -1
    ch = Mid$(args, i, 1)
    If ch < "0" Or ch > "9" Then Exit For
  Next i

  GetHwndFromCommand = CLng(Mid$(args, i + 1))
End Function

