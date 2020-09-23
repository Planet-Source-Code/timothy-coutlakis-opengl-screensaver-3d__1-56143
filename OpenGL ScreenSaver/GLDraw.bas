Attribute VB_Name = "GLDraw"
Option Explicit

Public Function DrawGLScene() As Boolean
' Here's Where We Do All The Drawing
  Dim r As Integer: r = RotAxis
  
  glClear clrColorBufferBit Or clrDepthBufferBit  ' Clear The Screen And The Depth Buffer
  glLoadIdentity                                  ' Reset The Current Modelview Matrix
    
  glLightfv ltLight0, lpmPosition, 0
  glLightfv ltLight1, lpmPosition, 0
  
  glTranslatef 0, 0, -5                  ' Move Ten Units Into The Screen

  If r = 1 Or r = 4 Then glRotatef rot, 1, 0, 0
  If r = 2 Or r = 4 Then glRotatef rot * 1.5, 0, 1, 0
  If r = 3 Or r = 4 Then glRotatef rot * 1.4, 0, 0, 1


  ' Pulsing Colors Based On The Rotation
  glColor3f 1 * Cos(rot / 200), 1 * Sin(rot / 250), 1 - 0.5 * Sin(rot / 170)

  If sTime Then
    glPrint Format(Time, TimeFormat)         ' Print GL Text To The Screen
  Else
    glPrint sText
  End If

  rot = rot + 0.5                     ' Increase The Rotation Variable
  DrawGLScene = True                              ' Everything Went OK
End Function

