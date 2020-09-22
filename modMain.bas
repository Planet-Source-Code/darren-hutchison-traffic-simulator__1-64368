Attribute VB_Name = "modMain"
Public Sub Main()

SetupGame
MainLoop
Terminate
    
End Sub

Public Sub MainLoop()

Dim lastTick As Long, booStartTimer As Boolean

Do
    booStartTimer = False
    lastTick = GetTickCount
    Do While (GetTickCount - lastTick) < FrameDelay
        DoEvents
    Loop
    
    If booAutoSpawn = True Then
        intSpawnCount = intSpawnCount + 1
        If intSpawnCount > 75 + Int(Rnd(50)) And intCarsOut < UBound(car) + 1 Then SpawnCar
    End If
    
    CheckKeys
    UpdateCarAI
    
    While booPause = True
        frmMain.Timer1.Enabled = False
        BltScreen
        CheckKeys
        DoEvents
        booStartTimer = True
    Wend
    
    If booStartTimer = True Then frmMain.Timer1.Enabled = True
    BltScreen
Loop While GameOver = False

End Sub

Public Sub DisplayLights()

Dim srcRect As RECT, dstRect As RECT, i As Integer

For i = 0 To 3
    Select Case i
        Case 0
            With dstRect
                .Top = 365
                .Left = 478
            End With
            srcRect.Left = 26 * intLightStatus(0)
        Case 1
            With dstRect
                .Top = 637
                .Left = 752
            End With
            srcRect.Left = 26 * intLightStatus(0)
        Case 2
            With dstRect
                .Top = 365
                .Left = 752
            End With
            srcRect.Left = 26 * intLightStatus(1)
        Case 3
            With dstRect
                .Top = 637
                .Left = 478
            End With
            srcRect.Left = 26 * intLightStatus(1)
    End Select
    
    srcRect.Top = 0: srcRect.Right = srcRect.Left + 25: srcRect.Bottom = 32
    With dstRect
        .Bottom = .Top + 32
        .Right = .Left + 25
    End With
    bltToBack dstRect, srfLights, srcRect, DDBLT_KEYSRC
Next i


End Sub

Public Sub UpdateLights()

Dim srcRect As RECT, dstRect As RECT

If booChange = False Then DisplayLights: Exit Sub

If booLightSetOneActive = True Then
    Select Case intLightStatus(0)
            Case 1
            intLightStatus(0) = 3
            Case 3
            intLightStatus(0) = 2
            Case 2
            booLightSetOneActive = False: intLightStatus(1) = 0
            Case 0
            intLightStatus(0) = 1
    End Select
Else
            Select Case intLightStatus(1)
            Case 1
            intLightStatus(1) = 3
            Case 3
            intLightStatus(1) = 2
            Case 2
            booLightSetOneActive = True: intLightStatus(0) = 0
            Case 0
            intLightStatus(1) = 1
    End Select
End If

booChange = False
frmMain.Timer1.Enabled = False
If intLightStatus(0) = 1 Or intLightStatus(1) = 1 Then frmMain.Timer1.Interval = 10000 Else frmMain.Timer1.Interval = 2000
frmMain.Timer1.Enabled = True
DisplayLights

End Sub

Public Sub SpawnCar()

Dim intStartArea As Integer, intLightSet As Integer

intCurrentCar = -1

Do
    intCurrentCar = intCurrentCar + 1
Loop Until car(intCurrentCar).booVisible = False

Do
    intStartArea = Int(Rnd * 4)
Loop Until intQuadrantLeftQueue(intStartArea) < 3 And intQuadrantRightQueue(intStartArea) < 3

If intStartArea = 0 Then car(intCurrentCar).Angle = 180
If intStartArea = 1 Then car(intCurrentCar).Angle = 270
If intStartArea = 2 Then car(intCurrentCar).Angle = 0
If intStartArea = 3 Then car(intCurrentCar).Angle = 90

If intLightStatus(CheckLightSet(car(intCurrentCar).intQuadrant)) = 3 Or intLightStatus(CheckLightSet(car(intCurrentCar).intQuadrant)) = 2 Then car(intCurrentCar).Speed = 1 Else car(intCurrentCar).Speed = 2

car(intCurrentCar).intQuadrant = intStartArea
car(intCurrentCar).intDirectionChosen = Int(Rnd * 3) '0=straight ahead, 1=left, 2=right
car(intCurrentCar).booVisible = True
car(intCurrentCar).x = intStartX(intStartArea)
car(intCurrentCar).y = intStartY(intStartArea)
car(intCurrentCar).booTurnState = False
car(intCurrentCar).booTurnAlternate = False
car(intCurrentCar).intLaneCounter = 0
car(intCurrentCar).intTurnCounter = 0
car(intCurrentCar).intStage = 1
car(intCurrentCar).intSpeedState = 0
car(intCurrentCar).intSpeedChangeCounter = 0
car(intCurrentCar).intSpeedChangeLimit = 90
car(intCurrentCar).Speed = 2
car(intCurrentCar).booReachedJunction = False
car(intCurrentCar).booLaneChanged = False
car(intCurrentCar).booCollision = False
car(intCurrentCar).booSlowedDown = False
car(intCurrentCar).booSpedUp = False
car(intCurrentCar).booIndicatorsOn = False
car(intCurrentCar).booIndicatorFlash = True
car(intCurrentCar).intIndicatorCounter = 0
intLightSet = CheckLightSet(car(intCurrentCar).intQuadrant)
If intLightStatus(intLightSet) > 1 Then
    If QuadrantQueueNumber(intCurrentCar) > 2 Then car(intCurrentCar).intSpeedState = 2
Else
    If QuadrantQueueNumber(intCurrentCar) > 2 Then car(intCurrentCar).Speed = 1: car(intCurrentCar).intSpeedState = 1
End If
intCarsOut = intCarsOut + 1
Select Case car(intCurrentCar).intDirectionChosen
    Case 2
    intQuadrantRightQueue(car(intCurrentCar).intQuadrant) = intQuadrantRightQueue(car(intCurrentCar).intQuadrant) + 1
    car(intCurrentCar).intLanePosition = intQuadrantRightQueue(car(intCurrentCar).intQuadrant)
    Case Else
    intQuadrantLeftQueue(car(intCurrentCar).intQuadrant) = intQuadrantLeftQueue(car(intCurrentCar).intQuadrant) + 1
    car(intCurrentCar).intLanePosition = intQuadrantLeftQueue(car(intCurrentCar).intQuadrant)
End Select
intSpawnCount = 0

End Sub

Public Sub CheckKeys()

    Static lastHelp As Long, lastPause As Long, lastSpawn As Long, lastAuto As Long
    
    If GetAsyncKeyState(vbKeyEscape) Then GameOver = True
    
    If GetAsyncKeyState(vbKeyF1) Then
        If GetTickCount - lastHelp > 1000 Then DisplayHelp = Not (DisplayHelp): lastHelp = GetTickCount
    End If
    
    If GetAsyncKeyState(vbKeyP) Then
        If GetTickCount - lastPause > 1000 Then
            If booPause = False Then booPause = True Else booPause = False
            lastPause = GetTickCount
        End If
    End If
    
    If GetAsyncKeyState(vbKeyS) Then
        If GetTickCount - lastSpawn > 1000 Then
            SpawnCar
            lastSpawn = GetTickCount
        End If
    End If
    
    If GetAsyncKeyState(vbKeyA) Then
        If GetTickCount - lastAuto > 1000 Then
            If booAutoSpawn = True Then booAutoSpawn = False Else booAutoSpawn = True
            lastAuto = GetTickCount
        End If
    End If
    
    If GetAsyncKeyState(vbKeyR) Then ResetProgram
    
End Sub

Public Sub BltScreen()
    
Dim srcRect As RECT, dstRect As RECT, size As Long, i As Integer
Static lastFPS As Long, tempFPS As Integer, lastAngle As Single

If (GetTickCount - lastFPS) > 1000 Then
    FPS = tempFPS: tempFPS = 0: lastFPS = GetTickCount
End If

tempFPS = tempFPS + 1
srcRect.Right = 1280: srcRect.Bottom = 1024
bltToBack srcRect, srfBackGround, srcRect, DDBLT_WAIT
BltText

For i = 0 To UBound(car)
    size = 71
    srcRect.Right = size: srcRect.Bottom = size
    dstRect.Left = car(i).x: dstRect.Right = car(i).x + size
    dstRect.Top = car(i).y: dstRect.Bottom = car(i).y + size
    If car(i).booVisible = True Then bltToBack dstRect, srfCarRotate(car(i).intColour, (car(i).Angle / 3)), srcRect, DDBLT_KEYSRC
    If car(i).booIndicatorsOn And car(i).intStage < 3 Then BltIndicators (i)
Next i

UpdateLights
modDirectX.flipPrimary

End Sub
Public Sub BltIndicators(intSelectedCar)

Dim srcRect As RECT, dstRect As RECT, size As Long, i As Integer, j As Integer

If car(intSelectedCar).intDirectionChosen = 0 Then Exit Sub

car(intSelectedCar).intIndicatorCounter = car(intSelectedCar).intIndicatorCounter + 1

If car(intSelectedCar).intIndicatorCounter >= 25 Then
    If car(intSelectedCar).booIndicatorFlash = True Then car(intSelectedCar).booIndicatorFlash = False Else car(intSelectedCar).booIndicatorFlash = True
    car(intSelectedCar).intIndicatorCounter = 0
End If

j = car(intSelectedCar).intDirectionChosen - 1
i = Int(car(intSelectedCar).Angle / 90) ' should be integer anyway

If car(intSelectedCar).booIndicatorFlash = True Then
    size = GetMaxRotateSize(srfIndicatorsL)
    srcRect.Right = size: srcRect.Bottom = size
    dstRect.Left = car(intSelectedCar).x: dstRect.Right = car(intSelectedCar).x + size
    dstRect.Top = car(intSelectedCar).y: dstRect.Bottom = car(intSelectedCar).y + size
    bltToBack dstRect, srfIndRotate(j, i), srcRect, DDBLT_KEYSRC
End If

End Sub
Public Sub BltText()

Dim i As Integer, j As Integer

If booPause = True Then textToBack 1000, 800, "Pause Active: " & booPause
'textToBack 1000, 700, "Junction Counter: " & intJunctionCounter
'textToBack 1000, 715, "AutoSpawn: " & booAutoSpawn

If DisplayHelp Then
    For i = 0 To 15
        If i > 7 Then j = 1 Else j = 0
        If car(i).booVisible = True Then
            textToBack car(i).x + 50, 40 + car(i).y + 50, "Distance to Junction: " & Abs(CartoJunctionDist(car(i).intQuadrant, car(i).x, car(i).y))
            textToBack car(i).x + 50, 55 + car(i).y + 50, "Stage: " & car(i).intStage
            textToBack car(i).x + 50, 70 + car(i).y + 50, "Quadrant: " & car(i).intQuadrant
            textToBack car(i).x + 50, 25 + car(i).y + 50, "Speed State: " & car(i).intSpeedState
            textToBack car(i).x + 50, 100 + car(i).y + 50, "Speed: " & car(i).Speed
            textToBack car(i).x + 50, 85 + car(i).y + 50, "SlowedDown sub: " & car(i).booSlowedDown
            textToBack car(i).x + 50, 115 + car(i).y + 50, "SpedUp sub: " & car(i).booSpedUp
        End If
    Next i
    
    textToBack 800, 100, "Total cars out: " & intCarsOut
    textToBack 800, 50, "Quadrant L Queue 0:" & intQuadrantLeftQueue(0)
    textToBack 800, 800, "Quadrant L Queue 1:" & intQuadrantLeftQueue(1)
    textToBack 300, 800, "Quadrant L Queue 2:" & intQuadrantLeftQueue(2)
    textToBack 300, 50, "Quadrant L Queue 3:" & intQuadrantLeftQueue(3)
    textToBack 800, 70, "Quadrant R Queue 0:" & intQuadrantRightQueue(0)
    textToBack 800, 820, "Quadrant R Queue 1:" & intQuadrantRightQueue(1)
    textToBack 300, 820, "Quadrant R Queue 2:" & intQuadrantRightQueue(2)
    textToBack 300, 70, "Quadrant R Queue 3:" & intQuadrantRightQueue(3)
End If

End Sub

Public Sub BltDelayText()

    textToBack 580, 450, "Loading:"
    
End Sub
