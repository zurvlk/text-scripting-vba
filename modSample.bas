Attribute VB_Name = "modSample"

Option Explicit


Sub dispOS()
  Dim nameOS As String
  nameOS = Application.OperatingSystem

  If nameOS Like "Windows *" Then
    ' Windows (32-bit) NT 6.01
    MsgBox "This is Windows OS. [" & nameOS & "]"

  ElseIf nameOS Like "Macintosh *" Then
    ' Macintosh (Intel) 10.8
    MsgBox "This is Mac OS X. [" & nameOS & "]"

  End If
End Sub


Sub dispVersion()
  MsgBox Application.Version
End Sub

' ëfêîÇ≈Ç‡îªíËÇ≥ÇπÇƒÇ®Ç≠
Public Function isPrime(ByVal num As Long) As Boolean
    If num < 2 Then
        isPrime = False
        Exit Function
    ElseIf num = 2 Then
        isPrime = True
        Exit Function
    ElseIf num Mod 2 = 0 Then
        isPrime = False
        Exit Function
    End If

    Dim sqrtNum As Double: sqrtNum = Sqr(num)
    Dim i As Integer
    For i = 3 To sqrtNum Step 2
        If num Mod i = 0 Then
            isPrime = False
            Exit Function
        End If
    Next i
    isPrime = True
End Function

' Ç®ÇΩÇ≠ÇÕâ~é¸ó¶ÇãÅÇﬂÇ™Çø
Public Function Gauss_Legendre(ByVal count As Integer) As Variant
    '// add declarations
    On Error GoTo catchError
    Dim a As Variant: a = 1
    Dim b As Variant: b = CDec(1 / Sqr(2))
    Dim t As Variant: t = 1 / 4
    Dim p As Variant: p = 1

    Dim preA As Variant
    Dim preB As Variant

    Dim i As Integer

    For i = 1 To count
        preA = CDec(a)
        preB = CDec(b)

        a = CDec((preA + preB) / 2)
        b = CDec(Sqr(preA * preB))
        t = CDec(t - p * (preA - a) ^ 2)
        p = CDec(2 * p)
    Next i

    Gauss_Legendre = CDec((a + b) ^ 2 / (4 * t))

exitFunction:
    Exit Function
catchError:
    '// add error handling
    GoTo exitFunction
End Function

Public Sub showPi()
    MsgBox Gauss_Legendre(5)
End Sub
