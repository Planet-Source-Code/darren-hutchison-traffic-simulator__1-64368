Attribute VB_Name = "modSetup"
'Traffic Simulator
'Darren Hutchison 2003
'www.dhutchison.freeuk.com
'Credit to Jim Camel for his DX Rotation code - it provided the inspiration and code base for this project

' API Functions for program timing, reading the keyboard and hiding the cursor
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'Setup constants for navigating the map
'Array values 0=top 1=right 2=bottom 3=left, similar to angles
Public intStartX(3) As Integer
Public intStop(3) As Integer
Public intStartY(3) As Integer

'Traffic Light variables
Public booLightSetOneActive As Boolean ' this is to indicate which set of traffic lights are changing their light state
Public booChange As Boolean
Public intLightStatus(1) As Integer    ' First element (0) is top left and bottom right light set
                                        ' Second element (1) is top right and bottom left light set
                                        ' 0==red/amber, 1=green, 2=red 3=amber for each traffic light set

'Car Variables
Public booAllCarsOut As Boolean ' don't spawn too many cars
Public intCarsOut As Integer ' counter of cars out
Public intSpawnCount As Integer ' used to provide gaps between cars spawning
Public booAutoSpawn As Boolean ' automatic/manual spawning of cars

'Global variables
Public intQuadrantLeftQueue(3) As Integer '0=top road 1=right road 2=bottom road 3=left road
Public intQuadrantRightQueue(3) As Integer '0=top road 1=right road 2=bottom road 3=left road
Public booJunctionQueue(3) As Boolean 'True=cars waiting to turn, False=clear
Public intJunctionCounter As Integer ' how many cars are currently inside the junction?
Public car(23) As tCar ' 24 cars maximum
Public GameOver As Boolean ' used to end the program
Public DisplayHelp As Boolean ' this may be removed on release
Public booPause As Boolean ' used for debugging
Public intCurrentCar As Integer ' which car is active?
Public FPS As Integer 'frames per second indicator
Public Const FrameDelay As Integer = 20 ' used to moderate program timing

'DirectX surface and other objects
Public srfBackGround As DirectDrawSurface7
Public srfTileset As DirectDrawSurface7
Public srfCar(5) As DirectDrawSurface7
Public srfCarRotate(5, 120) As DirectDrawSurface7
Public srfLights As DirectDrawSurface7
Public srfTitle As DirectDrawSurface7
Public srfIndicatorsL As DirectDrawSurface7
Public srfIndicatorsR As DirectDrawSurface7
Public srfIndRotate(1, 3) As DirectDrawSurface7
Public srfHeading As DirectDrawSurface7
Public ColorKey As DDCOLORKEY
Public fntCustom As New StdFont
Public Const PI As Double = 3.14159265358979

Public Sub SetupGame()
        
Dim i As Integer, j As Integer, booInitSuccess As Boolean

ShowCursor (0)
j = 0
Randomize
SetupMapVariables

intSpawnCount = 0
For i = 0 To UBound(car)
    car(i).Speed = 0
    car(i).booVisible = False
    car(i).intColour = j
    If j < 5 Then j = j + 1 Else j = 0
Next i

With fntCustom
    .name = "Trebuchet MS"
    .size = 22
    .Bold = False
End With

intCurrentCar = 0
booAllCarsOut = False
intCarsOut = 0
For i = 0 To 3
    booJunctionQueue(i) = False
    intQuadrantLeftQueue(i) = 0
    intQuadrantRightQueue(i) = 0
Next i

intLightStatus(0) = 1
intLightStatus(1) = 2
booChange = False
booLightSetOneActive = True
booPause = False
intJunctionCounter = 0
booAutoSpawn = True

booInitSuccess = initDirectDraw(frmMain, 1280, 1024, 16)

Set srfBackGround = loadDirectXSurface("", 1280, 1024)
Set srfTileset = loadDirectXSurface(App.Path + "\graphics\map3.jpg")
Set srfCar(0) = loadDirectXSurface(App.Path + "\graphics\car.gif")
Set srfCar(1) = loadDirectXSurface(App.Path + "\graphics\car2.gif")
Set srfCar(2) = loadDirectXSurface(App.Path + "\graphics\car3.gif")
Set srfCar(3) = loadDirectXSurface(App.Path + "\graphics\car4.gif")
Set srfCar(4) = loadDirectXSurface(App.Path + "\graphics\car5.gif")
Set srfCar(5) = loadDirectXSurface(App.Path + "\graphics\car6.gif")
Set srfIndicatorsL = loadDirectXSurface(App.Path + "\graphics\indicatorsl.gif")
Set srfIndicatorsR = loadDirectXSurface(App.Path + "\graphics\indicatorsr.gif")
Set srfLights = loadDirectXSurface(App.Path + "\graphics\lights.gif")
Set srfTitle = loadDirectXSurface(App.Path + "\graphics\title.gif")
Set srfHeading = loadDirectXSurface(App.Path + "\graphics\titletraffic.gif")
modDirectX.AddColorKey srfLights, ColorKey, vbWhite, vbWhite

BltInit
GenerateRotations
BltBackGround srfTileset, srfBackGround

With fntCustom
    .name = "Arial"
    .size = 10
    .Bold = False
End With

SetTextSize fntCustom
SetTexttoBack vbWhite

frmMain.Timer1.Enabled = True

End Sub
Public Sub BltInit()
    
Dim srcRect As RECT, dstRect As RECT, size As Long, i As Integer
Static lastFPS As Long, tempFPS As Integer, lastAngle As Single

If (GetTickCount - lastFPS) > 1000 Then
    FPS = tempFPS: tempFPS = 0: lastFPS = GetTickCount
End If

tempFPS = tempFPS + 1
srcRect.Right = 1280: srcRect.Bottom = 1024
SetTextSize fntCustom
SetTexttoBack vbBlue
With dstRect
    .Top = 500
    .Bottom = 550
    .Left = 500
    .Right = 550
End With
modDirectX.flipPrimary

End Sub
Public Sub SetupMapVariables()

'Top
intStartX(0) = 644
intStartY(0) = 0
intStop(0) = 320
'Right
intStartX(1) = 1206
intStartY(1) = 540
intStop(1) = 757
'Bottom
intStartX(2) = 545
intStartY(2) = 953
intStop(2) = 650
'Left
intStartX(3) = 0
intStartY(3) = 435
intStop(3) = 425

End Sub

Public Sub GenerateRotations()

Dim i As Integer, j As Integer, size As Long, lngFillvalue As Long
Dim Key As DDCOLORKEY, k As Integer
Dim srcRect As RECT, barRect As RECT, dstRect As RECT, lightRect As RECT, titleRectsrc As RECT, titleRectdest As RECT

'This sub generates the car rotations and displays a title screen with traffic light progress indicator
For j = 0 To 5
    size = 71 'GetMaxRotateSize(srfCar(j))
    Key.high = vbMagenta: Key.low = vbMagenta
    srcRect.Right = size: srcRect.Bottom = size
    barRect.Top = 0
    barRect.Left = 113
    barRect.Right = barRect.Left + 109
    barRect.Bottom = 332
    lightRect.Top = 0
    lightRect.Left = 0
    lightRect.Right = lightRect.Left + 109
    lightRect.Bottom = 332
    With titleRectsrc
        .Top = 0
        .Left = 0
        .Right = .Left + 307
        .Bottom = 200
    End With
    With titleRectdest
        .Top = 80
        .Left = 410
        .Right = .Left + 460
        .Bottom = .Top + 300
    End With
    For i = 0 To 120
        Set srfCarRotate(j, i) = loadDirectXSurface("", size, size)
        srfCarRotate(j, i).SetColorKey DDCKEY_SRCBLT, Key
        srfCarRotate(j, i).BltColorFill srcRect, vbRed
        rotateSurface srfCar(j), srfCarRotate(j, i), i * 3, size / 4, 0, 0
        bltToBack titleRectdest, srfHeading, titleRectsrc, DDBLT_WAIT
        BltDelayText
        With dstRect
            .Top = 500
            .Left = 580
            .Right = .Left + 109
            .Bottom = .Top + 332
        End With
        k = j + 1
        bltToBack dstRect, srfTitle, barRect, DDBLT_KEYSRC
        If j <= 2 Then lightRect.Bottom = 111: dstRect.Bottom = dstRect.Top + 111
        If j < 5 And j > 2 Then lightRect.Bottom = 222: dstRect.Bottom = 222 + dstRect.Top
        If j >= 5 Then lightRect.Bottom = 332: dstRect.Bottom = dstRect.Top + 333: dstRect.Top = dstRect.Bottom - 111: lightRect.Top = 222
        lightRect.Left = 0: lightRect.Right = 109
        bltToBack dstRect, srfTitle, lightRect, DDBLT_KEYSRC
        modDirectX.flipPrimary
        DoEvents
    Next i
Next j

srcRect.Top = 0
srcRect.Left = 0
'Rotations for the indicators - added to end - shouldn't delay the load routine much
For j = 0 To 1
    For i = 0 To 3
        size = GetMaxRotateSize(srfIndicatorsL)
        Key.high = vbMagenta: Key.low = vbMagenta
        srcRect.Right = size: srcRect.Bottom = size
        Set srfIndRotate(j, i) = loadDirectXSurface("", size, size)
        srfIndRotate(j, i).SetColorKey DDCKEY_SRCBLT, Key
        srfIndRotate(j, i).BltColorFill srcRect, vbRed
        If j = 1 Then
            rotateSurface srfIndicatorsL, srfIndRotate(j, i), i * 90, size / 4, 0, 0
        Else
            rotateSurface srfIndicatorsR, srfIndRotate(j, i), i * 90, size / 4, 0, 0
        End If
    Next i
Next j

End Sub

Public Sub BltBackGround(ByRef TileSet As DirectDrawSurface7, ByRef dest As DirectDrawSurface7)

Dim srcRect As RECT, dstRect As RECT

With srcRect
    .Left = 0: .Top = 0
    .Bottom = .Top + 1024: .Right = .Left + 1280
End With

With dstRect
    .Left = 0: .Top = 0
    .Bottom = .Top + 1024: .Right = .Left + 1280
End With
       
dest.Blt dstRect, TileSet, srcRect, DDBLT_WAIT

End Sub

Public Sub Terminate()

Dim i As Integer, j As Integer

For j = 0 To 5
    Set srfCar(j) = Nothing
    
    For i = 0 To 120
        Set srfCarRotate(j, i) = Nothing
    Next i
Next j

ShowCursor (1)
Set srfTileset = Nothing
Set srfBackGround = Nothing
modDirectX.terminateDirectX
Unload frmMain

End Sub

Public Sub ResetProgram()

Dim i As Integer

intSpawnCount = 0
For i = 0 To UBound(car)
    car(i).Speed = 0
    car(i).booVisible = False
    car(i).intColour = j
    If j < 5 Then j = j + 1 Else j = 0
Next i

With fntCustom
    .name = "Trebuchet MS"
    .size = 22
    .Bold = False
End With

intCurrentCar = 0
booAllCarsOut = False
intCarsOut = 0
For i = 0 To 3
    booJunctionQueue(i) = False
    intQuadrantLeftQueue(i) = 0
    intQuadrantRightQueue(i) = 0
Next i

intLightStatus(0) = 1
intLightStatus(1) = 2
booChange = False
booLightSetOneActive = True
booPause = False
intJunctionCounter = 0

End Sub
