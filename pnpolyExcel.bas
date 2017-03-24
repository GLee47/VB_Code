Public Function pnpoly(nvert As Integer, vertx1 As Double, verty1 As Double, vertx2 As Double, verty2 As Double, testx As Double, testy As Double) As Integer

    If (((verty2 > testy) <> (verty1 > testy)) And (testx < (vertx1 - vertx2) * (testy - verty2) / (verty1 - verty2) + vertx2)) Then
        pnpoly = 1
        Exit Function
    Else
        pnpoly = 0
    End If
End Function
