Option Explicit

Function cat(r As Range)
    Dim i As Variant
    For Each i In r
       cat = cat & i
    Next i
End Function

Function bigCat(r1 As Range, Optional r2 As Range, Optional r3 As Range)
    Dim i, j, k As Variant
    For Each i In r1
       bigCat = bigCat & i
    Next i
    
    If Not r2 Is Nothing Then
        For Each j In r2
            bigCat = bigCat & j
        Next j
    End If
    
    If Not r3 Is Nothing Then
        For Each k In r3
            bigCat = bigCat & k
        Next k
    End If
End Function

Function bossCat(ParamArray r())
    Dim ub As Integer
    ub = UBound(r)

    Dim i As Integer
    Dim k As Variant
    
    For i = 0 To ub
        For Each k In r(i)
            bossCat = bossCat & k
        Next k
    Next i

End Function


Function dog(r As Variant)
    If TypeName(r) = "Range" Then
        Dim i As Variant
        For Each i In r
            dog = dog & i
        Next i
    Else
        dog = r
    End If
    
End Function

Function bigDog(r1 As Variant, Optional r2 As Variant = "", Optional r3 As Variant = "")
    Dim i, j, k As Variant
    If TypeName(r1) = "Range" Then
        For Each i In r1
           bigDog = bigDog & i
        Next i
    Else
        bigDog = r1
    End If
    
    If Not IsEmpty(r2) Then
        If TypeName(r2) = "Range" Then
            For Each j In r2
               bigDog = bigDog & j
            Next j
        Else
            bigDog = bigDog & r2
        End If
    End If
    
    If Not IsEmpty(r3) Then
        If TypeName(r3) = "Range" Then
            For Each k In r3
               bigDog = bigDog & k
            Next k
        Else
            bigDog = bigDog & r3
        End If
    End If

End Function

Function bossDog(ParamArray r())
    Dim ub As Integer
    ub = UBound(r)

    Dim i As Integer
    Dim k As Variant
    
    For i = 0 To ub
        MsgBox TypeName(r(i))
        If TypeName(r(i)) = "Range" Then
            For Each k In r(i)
                bossDog = bossDog & k
            Next k
        Else
            bossDog = bossDog & r(i)
        End If
    Next i

End Function

Sub j()
    Dim a As Range
    MsgBox IsEmpty(a)
End Sub
