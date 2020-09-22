Attribute VB_Name = "modAI"
' Car variable declarations are in here to make AI development easier

Public Type tCar
    x As Single ' x co-ordinate of car
    y As Single ' y co-ordinate of car
    Speed As Long '0=stopped 1=slow 2=normal 3=fast
    Angle As Integer ' direction car travels - 0 to 270 degrees
    booVisible As Boolean ' car is spawned or not
    intColour As Integer ' choice of 6 car colours
    intDirectionChosen As Integer '0=straight ahead 1=left 2=right
    booTurnState As Boolean ' marker to indicate that turn is complete
    intTurnCounter As Integer ' counter to track angle change of car
    booTurnAlternate As Boolean ' this makes the car turn less sharply
    intStage As Integer '1=approaching junction 2=in junction 3=after junction
    intQuadrant As Integer '0=top road 1=right road 2=bottom road 3=left road
    intSpeedState As Integer ' 0=steady 1=speeding up 2=slowing down
    intSpeedChangeCounter As Integer ' used to control smooth changes in speed
    intSpeedChangeLimit As Integer ' used to control smooth changes in speed - lower=more drastic changes
    intLaneCounter As Integer ' counter to track movement into a lane
    intLanePosition As Integer ' counter to track where a car will stop
    booLaneChanged As Boolean ' marker to indicate that lane has been chosen
    booReachedJunction As Boolean ' subroutine complete flag
    booSlowedDown As Boolean ' subroutine complete flag
    booSpedUp As Boolean ' subroutine complete flag
    booCollision As Boolean ' subroutine complete flag
    booIndicatorsOn As Boolean ' Indicators on/off
    booIndicatorFlash As Boolean ' Flashing indicators
    intIndicatorCounter As Integer ' Timer to control flash on/off
End Type

Public Sub UpdateCarAI()

Dim intCars As Integer

For intCars = 0 To UBound(car)
    If car(intCars).booVisible = True Then
            If car(intCars).intStage = 1 Then
                ChooseLane (intCars)
                'CheckCollision (intCars)
                If car(intCars).booCollision = False Then
                    SlowDownAtLights (intCars)
                    SpeedUpAtLights (intCars)
                    ReachedJunction (intCars)
                End If
            End If
            If car(intCars).intStage = 2 Then TurnCar (intCars)
            If car(intCars).intStage = 3 Then SpeedUpCar (intCars)
            ModifyCarSpeed (intCars)
            UpdateCarCoords (intCars)
    End If
Next intCars

End Sub

Public Sub ChooseLane(intSelectedCar)

Dim i As Integer

If car(intSelectedCar).booLaneChanged = True Then Exit Sub

If car(intSelectedCar).intLaneCounter >= 50 Then
    car(intSelectedCar).booLaneChanged = True
    Exit Sub
End If

i = intSelectedCar
If Abs(CartoJunctionDist(car(intSelectedCar).intQuadrant, car(intSelectedCar).x, car(intSelectedCar).y)) > (LargerQueueValue(i, car(intSelectedCar).intQuadrant) * 90) + 150 Then Exit Sub

If car(intSelectedCar).intDirectionChosen <> 0 Then car(intSelectedCar).booIndicatorsOn = True

If car(intSelectedCar).intQuadrant = 0 Then
    Select Case car(intSelectedCar).intDirectionChosen
        Case 2
            car(intSelectedCar).x = car(intSelectedCar).x - 0.5
        Case Else
            car(intSelectedCar).x = car(intSelectedCar).x + 0.5
    End Select
End If

If car(intSelectedCar).intQuadrant = 1 Then
    Select Case car(intSelectedCar).intDirectionChosen
        Case 2
            car(intSelectedCar).y = car(intSelectedCar).y - 0.5
        Case Else
            car(intSelectedCar).y = car(intSelectedCar).y + 0.5
    End Select
End If

If car(intSelectedCar).intQuadrant = 2 Then
    Select Case car(intSelectedCar).intDirectionChosen
        Case 2
            car(intSelectedCar).x = car(intSelectedCar).x + 0.5
        Case Else
            car(intSelectedCar).x = car(intSelectedCar).x - 0.5
    End Select
End If

If car(intSelectedCar).intQuadrant = 3 Then
    Select Case car(intSelectedCar).intDirectionChosen
        Case 2
            car(intSelectedCar).y = car(intSelectedCar).y + 0.5
        Case Else
            car(intSelectedCar).y = car(intSelectedCar).y - 0.5
    End Select
End If

car(intSelectedCar).intLaneCounter = car(intSelectedCar).intLaneCounter + 1

End Sub

Public Sub CheckCollision(intSelectedCar)

Dim i As Integer, j As Integer

j = intSelectedCar

For i = 0 To UBound(car)
    If car(i).booVisible = True Then
        If car(i).intQuadrant = car(j).intQuadrant Then CollisionDistance j, car(i).x, car(i).y, car(j).x, car(j).y, car(j).Angle
    End If
Next i

End Sub
Public Sub SlowDownAtLights(intSelectedCar)

Dim intLightSet As Integer

If car(intSelectedCar).booSlowedDown = True Or car(intSelectedCar).booSpedUp = True Then Exit Sub

intLightSet = CheckLightSet(car(intSelectedCar).intQuadrant)
If intLightStatus(intLightSet) <> 1 Then
    If Abs(CartoJunctionDist(car(intSelectedCar).intQuadrant, car(intSelectedCar).x, car(intSelectedCar).y)) <= 90 Then
        car(intSelectedCar).intSpeedState = 1
        car(intSelectedCar).booSlowedDown = True
        Exit Sub
    End If
    If Abs(CartoJunctionDist(car(intSelectedCar).intQuadrant, car(intSelectedCar).x, car(intSelectedCar).y)) < (100 + ((car(intSelectedCar).intLanePosition - 1) * 90)) Then
        car(intSelectedCar).intSpeedState = 2
        car(intSelectedCar).booSlowedDown = True
        If car(intSelectedCar).booSpedUp = True Then car(intSelectedCar).booSpedUp = False
    End If
End If

End Sub

Public Sub ModifyCarSpeed(intSelectedCar)

If car(intSelectedCar).intSpeedState = 2 Then
    If car(intSelectedCar).Speed = 0 Then GoTo SetZeroSpeedValues
    If car(intSelectedCar).intSpeedChangeCounter < car(intSelectedCar).intSpeedChangeLimit Then
        car(intSelectedCar).intSpeedChangeCounter = car(intSelectedCar).intSpeedChangeCounter + 1
        If car(intSelectedCar).intSpeedChangeCounter = 1 Then car(intSelectedCar).Speed = car(intSelectedCar).Speed - 1
    Else
        GoTo SetZeroSpeedValues
    End If
ElseIf car(intSelectedCar).intSpeedState = 1 Then
    If car(intSelectedCar).Speed = 2 Then GoTo SetZeroSpeedValues
    If car(intSelectedCar).intSpeedChangeCounter < 20 Then
        car(intSelectedCar).intSpeedChangeCounter = car(intSelectedCar).intSpeedChangeCounter + 1
        If car(intSelectedCar).intSpeedChangeCounter = 1 Then car(intSelectedCar).Speed = car(intSelectedCar).Speed + 1
    Else
        GoTo SetZeroSpeedValues
    End If
End If
Exit Sub

SetZeroSpeedValues:

If car(intSelectedCar).Speed >= 2 Or car(intSelectedCar).Speed = 0 Then car(intSelectedCar).intSpeedState = 0
car(intSelectedCar).intSpeedChangeCounter = 0

End Sub
Public Sub SpeedUpAtLights(intSelectedCar)

Dim intLightSet As Integer

'If car(intSelectedCar).booSpedUp = True Then Exit Sub 'Or car(intSelectedCar).booSlowedDown = False Then Exit Sub
If car(intSelectedCar).Speed <> 0 Then Exit Sub
'If Abs(CartoJunctionDist(car(intSelectedCar).intQuadrant, car(intSelectedCar).x, car(intSelectedCar).y)) > 2 Then Exit Sub

intLightSet = CheckLightSet(car(intSelectedCar).intQuadrant)
If intLightStatus(intLightSet) = 1 Then car(intSelectedCar).intSpeedState = 1: car(intSelectedCar).booSpedUp = True

End Sub

Public Sub ReachedJunction(intSelectedCar)

If car(intSelectedCar).booReachedJunction = True Then Exit Sub

If Abs(CartoJunctionDist(car(intSelectedCar).intQuadrant, car(intSelectedCar).x, car(intSelectedCar).y)) <= 2 Then
    car(intSelectedCar).intStage = 2
    intJunctionCounter = intJunctionCounter + 1
    Select Case car(intSelectedCar).intDirectionChosen
    Case 2
        intQuadrantRightQueue(car(intSelectedCar).intQuadrant) = intQuadrantRightQueue(car(intSelectedCar).intQuadrant) - 1
    Case Else
        intQuadrantLeftQueue(car(intSelectedCar).intQuadrant) = intQuadrantLeftQueue(car(intSelectedCar).intQuadrant) - 1
    End Select
    car(intSelectedCar).booReachedJunction = True
End If

End Sub

Public Sub TurnCar(i As Integer)

Dim intTurnvalue As Integer

If car(i).intDirectionChosen = 0 Then
    car(i).intSpeedState = 1
    If car(i).intQuadrant < 2 Then car(i).intQuadrant = car(i).intQuadrant + 2 Else car(i).intQuadrant = car(i).intQuadrant - 2
    car(i).intStage = 3: intJunctionCounter = intJunctionCounter - 1
    Exit Sub
End If

If car(i).intTurnCounter >= 30 Then
    car(i).intTurnCounter = 0
    car(i).booTurnState = False
    car(i).intStage = 3
    intJunctionCounter = intJunctionCounter - 1
    If car(i).intDirectionChosen = 1 Then car(i).intQuadrant = car(i).intQuadrant + 1
    If car(i).intDirectionChosen = 2 Then car(i).intQuadrant = car(i).intQuadrant - 1
    If car(i).intQuadrant = -1 Then car(i).intQuadrant = 3
    If car(i).intQuadrant = 4 Then car(i).intQuadrant = 0
    car(i).intDirectionChosen = 0
    Exit Sub
End If

If car(i).intDirectionChosen = 2 Then intTurnvalue = 135
If car(i).intDirectionChosen = 1 Then intTurnvalue = 30

If Abs(CartoJunctionDist(car(i).intQuadrant, car(i).x, car(i).y)) < intTurnvalue And car(i).booTurnState = False Then Exit Sub
car(i).Speed = 2
car(i).booTurnState = True
car(i).booIndicatorsOn = False
If car(i).booTurnAlternate = True Then car(i).booTurnAlternate = False: Exit Sub
car(i).booTurnAlternate = True
If car(i).intDirectionChosen = 2 Then car(i).Angle = (car(i).Angle + 3) Mod 360
If car(i).intDirectionChosen = 1 Then car(i).Angle = (car(i).Angle - 3) Mod 360
car(i).intTurnCounter = car(i).intTurnCounter + 1
If car(i).Angle < 0 Then car(i).Angle = 360 + car(i).Angle

End Sub

Public Sub SpeedUpCar(i As Integer)

If Abs(CartoJunctionDist(car(i).intQuadrant, car(i).x, car(i).y)) = 50 Then car(i).Speed = 3: car(i).intSpeedState = 0

End Sub

Public Sub UpdateCarCoords(i As Integer)

Dim AngleRad As Single

AngleRad = car(i).Angle * PI / 180
Select Case car(i).Angle
    Case 0
        car(i).y = car(i).y - car(i).Speed
    Case 1 To 89
        car(i).y = car(i).y - (car(i).Speed * Cos(AngleRad))
        car(i).x = car(i).x + (car(i).Speed * Sin(AngleRad))
    Case 90
        car(i).x = car(i).x + car(i).Speed
    Case 91 To 179
        car(i).y = car(i).y + (car(i).Speed * Sin(AngleRad - 1.5707963267949))
        car(i).x = car(i).x + (car(i).Speed * Cos(AngleRad - 1.5707963267949))
    Case 180
        car(i).y = car(i).y + car(i).Speed
    Case 181 To 269
        car(i).y = car(i).y + (car(i).Speed * Cos(AngleRad - PI))
        car(i).x = car(i).x - (car(i).Speed * Sin(AngleRad - PI))
    Case 270
        car(i).x = car(i).x - car(i).Speed
    Case 271 To 359
        car(i).y = car(i).y - (car(i).Speed * Sin(AngleRad - 4.71238898038469))
        car(i).x = car(i).x - (car(i).Speed * Cos(AngleRad - 4.71238898038469))
End Select

If car(i).x > 1280 - 64 Or car(i).y > 1024 - 64 Or car(i).x < 0 Or car(i).y < 0 Then
    car(i).booVisible = False
    intCarsOut = intCarsOut - 1
End If

End Sub
