Attribute VB_Name = "modFunctions"
Public Function CartoJunctionDist(intQuadrant As Integer, sngCarXValue As Single, sngCarYValue As Single) As Integer

Dim intXorY As Integer, intStopValue As Integer

If intQuadrant = 2 Or intQuadrant = 0 Then intXorY = 1 Else intXorY = 0
If intXorY = 1 Then CartoJunctionDist = Int(sngCarYValue - intStop(intQuadrant))
If intXorY = 0 Then CartoJunctionDist = Int(sngCarXValue - intStop(intQuadrant))

End Function

Public Function CheckLightSet(i As Integer) As Integer

If i = 3 Or i = 1 Then CheckLightSet = 0 Else CheckLightSet = 1

End Function

Public Function OppositeQuad(i As Integer) As Integer

If i > 1 Then OppositeQuad = i - 2 Else OppositeQuad = i + 2

End Function

Public Function LargerQueueValue(intCar As Integer, intQuadrant As Integer) As Integer

If intQuadrantLeftQueue(car(intCar).intQuadrant) > intQuadrantRightQueue(car(intCar).intQuadrant) Then
    LargerQueueValue = intQuadrantLeftQueue(car(intCar).intQuadrant)
Else
    LargerQueueValue = intQuadrantRightQueue(car(intCar).intQuadrant)
End If

End Function

Public Function QuadrantQueueNumber(intCar As Integer) As Integer

If car(intCarNumber).intDirectionChosen = 2 Then
    QuadrantQueueNumber = intQuadrantRightQueue(car(intCar).intQuadrant)
Else
    QuadrantQueueNumber = intQuadrantLeftQueue(car(intCar).intQuadrant)
End If

End Function

Public Function CollisionDistance(intCar As Integer, sngFrontCarx As Single, sngFrontCary As Single, sngBackCarx As Single, sngBackCary As Single, intCarAngle As Integer) As Integer

Dim booXSignificant As Boolean, sngFront As Single, sngBack As Single


If sngFrontCarx = sngBackCarx Then
    booXSignificant = False
ElseIf sngFrontCary = sngBackCary Then booXSignificant = True
Else: car(intCar).booCollision = False: Exit Function
End If

If booXSignificant = True Then sngFront = sngFrontCarx: sngBack = sngBackCarx Else sngFront = sngFrontCary: sngBack = sngBackCary

car(intCar).booCollision = True

End Function
