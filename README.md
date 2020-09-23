<div align="center">

## Another Star Field \(updated\!\) \- well documented


</div>

### Description

Draws an animated StarField. A left-click with the mouse will move the StarField center to the mouse position, holding down the left mouse button while dragging the mouse will continually change the StarField center, holding down the right mouse button will activate a "hyperspace" effect (of sorts). The form can be resized. Each star's size and brightness is calculated according to its relative distance from you, the viewer. The number of stars in the StarField is easily changed.
 
### More Info
 
(1) Start a new project (2) Add a "timer" object to the existing form and set its "interval" property to "1" (3) Paste the supplied code into the Form code-window (4) Run it (5) Enjoy! (6) Notes: The vortex center can be changed by left-clicking with the mouse (or dragging the mouse with left button down) and the form can be resized, A HyperSpace effect (of sorts) can be activated by holding down the right mouse button, The number of stars can be changed by modifying the value of "gStarCount". "Submitting" this code seems to screw up its format (alignment and such), sorry....


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeff Godfrey](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeff-godfrey.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeff-godfrey-another-star-field-updated-well-documented__1-2328/archive/master.zip)





### Source Code

```
Option Explicit
' Define a Star
Private Type StarType
  xs As Long    ' X start coordinate
  ys As Long    ' Y start coordinate
  xe As Long    ' X end coordinate
  ye As Long    ' Y end coordinate
  Speed As Single  ' Star speed
End Type
'Number of Stars in the StarField
Const gStarCount = 150
' Define a "StarField" as a certain number of "Stars"
Dim StarField(gStarCount) As StarType
Dim gXCen As Long     ' x center of vortex
Dim gYCen As Long     ' y center of vortex
Dim gXVortexLow As Long  ' left most edge of vortex
Dim gXVortexHigh As Long  ' right most edge of vortex
Dim gYVortexLow As Long  ' top edge of vortex
Dim gYVortexHigh As Long  ' bottom edge of vortex
Dim gMaxRad As Long    ' used to adjust star "brightness"
Dim gHyperSpace As Boolean ' used to toggle hyperspace mode
Private Sub Form_Load()
  ' assign several Form properties
  Me.BackColor = vbBlack
  Me.Caption = "StarField - Jeff Godfrey"
  Me.Show
  Me.WindowState = vbMaximized
  ' assign vortex center to be the form center
  GetNewVortex Me.ScaleWidth / 2, Me.ScaleHeight / 2
  ' initialize all Star objects
  InitStars
End Sub
' initialize all Star objects
Sub InitStars()
  Dim i As Integer
  For i = 1 To gStarCount
    ' assign locations and speeds to all Stars in the StarField
    StarField(i).xs = (gXVortexHigh - gXVortexLow - 1) * Rnd + gXVortexLow
    StarField(i).ys = (gYVortexHigh - gYVortexLow - 1) * Rnd + gYVortexLow
    StarField(i).xe = StarField(i).xs
    StarField(i).ye = StarField(i).ys
    StarField(i).Speed = Rnd + 0.1   ' (.1 - 1.1)
  Next i
End Sub
' if the left mouse button was clicked, reassign vortex center
' to mouse location...
' if the right mouse button was clicked, activate
' "hyperspace" mode
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (Button = vbLeftButton) Then
    GetNewVortex X, Y
  ElseIf (Button = vbRightButton) Then
    gHyperSpace = True
  End If
End Sub
' If the mouse is moved with the left button held down,
' continually change the vortex center
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (Button = vbLeftButton) Then
    GetNewVortex X, Y
  End If
End Sub
' if the right button was just released...
' deactivate hyperspace mode and erase the hyperspace effect
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (Button = vbRightButton) Then
    gHyperSpace = False
    Me.Cls
  End If
End Sub
' if the form is resized, reassign the vortex center to the new window center
Private Sub Form_Resize()
  ' recalculate new vortex information based on current form dimensions
  GetNewVortex Me.ScaleWidth / 2, Me.ScaleHeight / 2
  ' if window is minimized or maximized, don't resize it
  ' (this will prevent a RunTime error...)
  If (Me.WindowState = vbMaximized) Then Exit Sub
  If (Me.WindowState = vbMinimized) Then Exit Sub
  ' ensure form is not made too small - this will
  ' prevent a possible "divide by 0" error...
  If Me.Width < 500 Then Me.Width = 500
  If Me.Height < 1500 Then Me.Height = 1500
End Sub
' Assign new vortex and other misc variables
' input: The X,Y coordinates of the new vortex center
' output: Nothing (reassigns global vortex variables)
Sub GetNewVortex(ByVal VortexgXCen As Long, ByVal VortexgYCen As Long)
  Dim XOffset As Long ' a +/- X range from the vortex center
  Dim YOffset As Long ' a +/- Y range from the vortex center
  gXCen = VortexgXCen  ' the GLOBAL center of the vortex
  gYCen = VortexgYCen  ' the GLOBAL center of the vortex
  ' calculate a range distance from the vortex center.
  XOffset = Int(Me.Width * 0.1)
  YOffset = Int(Me.Height * 0.1)
  ' calculate the GLOBAL actual range for both axis'
  ' a new star will always be "born" within this area...
  gXVortexLow = gXCen - XOffset
  gXVortexHigh = gXCen + XOffset
  gYVortexLow = gYCen - YOffset
  gYVortexHigh = gYCen + YOffset
  ' Assign a GLOBAL "maximum screen radius". This is
  ' used in the Star's brightness calculation
  If (Me.ScaleWidth < Me.ScaleHeight) Then
    gMaxRad = Int(Me.ScaleWidth / 2)
  Else
    gMaxRad = Int(Me.ScaleHeight / 2)
  End If
End Sub
' when the timer fires, animate each Star in the StarField
' this is where all the interesting stuff happens...
Private Sub Timer1_Timer()
  Dim i As Integer
  Dim XVector As Long    ' current Star's X distance from "vortex" center
  Dim YVector As Long    ' current Star's Y distance from "vortex" center
  Dim NewXe As Long     ' New X end coord of current Star
  Dim NewYe As Long     ' New Y end coord of current Star
  Dim NewXs As Long     ' New X start coord of current Star
  Dim NewYs As Long     ' New Y start coord of current Star
  Dim Speed As Single    ' Speed of current Star
  Dim Range As Integer   ' Range of current Star
  Dim DrawColor As Integer ' Color of current Star
  Dim EraseColor As Integer ' Erase color (the form's background color)
  ' assign the erase color to be the form background color
  EraseColor = Me.BackColor
  ' for each Star in the StarField...
  For i = 1 To gStarCount
    ' set new startpoint equal to the Star's previous endpoint
    NewXs = StarField(i).xe
    NewYs = StarField(i).ye
    Speed = StarField(i).Speed
    ' calculate X and Y distances from the current "vortex" center
    XVector = Abs(gXCen - NewXs)
    YVector = Abs(gYCen - NewYs)
    ' calculate Star's X direction and length based on current "vortex" X center
    If (NewXs > gXCen) Then
      NewXe = NewXs + Int(XVector * 0.2) * Speed
    Else
      NewXe = NewXs - Int(XVector * 0.2) * Speed
    End If
    ' calcuate Star's Y direction and length based on current "vortex" Y center
    If (NewYs > gYCen) Then
      NewYe = NewYs + Int(YVector * 0.2) * Speed
    Else
      NewYe = NewYs - Int(YVector * 0.2) * Speed
    End If
    ' if not in hyperspace mode...
    ' erase previous copy of the current Star (draw in backcolor)
    If (Not gHyperSpace) Then
      Me.Line (StarField(i).xs, StarField(i).ys)- _
          (StarField(i).xe, StarField(i).ye), EraseColor
    End If
    ' if new start coord is off the screen, reset it "near" the "vortex" center
    If (NewXs < 0 Or NewXs > Me.ScaleWidth Or _
      NewYs < 0 Or NewYs > Me.ScaleHeight) Then
      StarField(i).xs = (gXVortexHigh - gXVortexLow - 1) * Rnd + gXVortexLow
      StarField(i).ys = (gYVortexHigh - gYVortexLow - 1) * Rnd + gYVortexLow
      StarField(i).xe = StarField(i).xs
      StarField(i).ye = StarField(i).ys
    ' if new start coord is on the screen, draw new Star vector
    Else
      ' see how far the Star is from the "vortex" center
      ' this is used to determine its "brightness"...
      Range = GetStarRange(NewXs, NewYs)
      DrawColor = Range * 25
      ' draw the Star at its new location
      ' the Star color can be changed here (currently yellow...)
      Me.Line (NewXs, NewYs)-(NewXe, NewYe), RGB(DrawColor, DrawColor, 0)
      ' store Star endpoints for next erase cycle...
      StarField(i).xs = NewXs
      StarField(i).ys = NewYs
      StarField(i).xe = NewXe
      StarField(i).ye = NewYe
    End If
  Next i
End Sub
' determine how far the Star is from the "vertex" center
' used to determine the Star's brightness
' Note: Since this routine is called within the main animation
'    loop, it is VERY EXPENSIVE (in CPU cycles) due the
'    muliply, divide, and square root math. There should
'    be a better way, but this will work for now...
' Input: X and Y coordinate of current star
' Output: An integer in the range of 1-10
Function GetStarRange(ByVal X As Long, ByVal Y As Long) As Integer
  Dim Dist As Long
  Dim XVector As Long
  Dim YVector As Long
  XVector = Abs(gXCen - X)
  YVector = Abs(gYCen - Y)
  ' Calculate distance from "vortex" center
  Dist = Sqr(XVector * XVector + YVector * YVector)
  ' return value in the range of 1-10
  GetStarRange = Int((Dist / gMaxRad) * 10)
  If (GetStarRange < 1) Then GetStarRange = 1
  If (GetStarRange > 10) Then GetStarRange = 10
End Function
```

