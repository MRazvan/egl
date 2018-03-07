Attribute VB_Name = "Utils"
Function Clamp(val As Double, min As Double, max As Double) As Double
    Clamp = WorksheetFunction.max(min, WorksheetFunction.min(max, val))
End Function

Public Function GetProcAddress(ByVal lngAddressOf As Long) As Long
  GetProcAddress = lngAddressOf
End Function

Public Sub Swap(ByRef a As Variant, ByRef b As Variant)
    Dim tmp As Variant
    tmp = a
    a = b
    b = tmp
End Sub

Public Sub ClearMatrix(ByRef arr As Matrix)
    Dim x As Integer
    Dim y As Integer
    For x = 0 To 3
        For y = 0 To 3
            arr.MatVal(x, y) = 0#
        Next y
    Next x
End Sub

Public Sub IdentMatrix(ByRef arr As Matrix)
    ClearMatrix arr
    arr.MatVal(0, 0) = 1#
    arr.MatVal(1, 1) = 1#
    arr.MatVal(2, 2) = 1#
    arr.MatVal(3, 3) = 1#
End Sub

Public Function MatMult3(ByRef m1 As Matrix, ByRef m2 As Matrix, ByRef m3 As Matrix) As Matrix
    Dim res As New Matrix
    res.Data = WorksheetFunction.MMult(m1.Data, WorksheetFunction.MMult(m2.Data, m3.Data))
    Set MatMult3 = res
End Function

Public Function MP_Mult(ByRef modelview As Matrix, ByRef projMat As Matrix) As Matrix
    Dim res As New Matrix
    res.Data = WorksheetFunction.MMult(modelview.Data, projMat.Data)
    Set MP_Mult = res
End Function

Public Sub VectorMatrixMult(ByRef mat As Matrix, ByRef v As Vec4)
    Dim vec
    vec = Array(v.x, v.y, v.z, v.w)
    vec = WorksheetFunction.MMult(vec, mat.Data)
    v.x = vec(1)
    v.y = vec(2)
    v.z = vec(3)
    v.w = vec(4)
End Sub

Public Function Map(val As Double, minValRange As Double, maxValRange As Double, targetRangeMin As Double, targetRangeMax As Double)
    Map = (val - minValRange) * (targetRangeMax - targetRangeMin) / (maxValRange - minValRange) + targetRangeMin
End Function
