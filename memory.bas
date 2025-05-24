Attribute VB_Name = "Module1"
Sub Memory()
    Dim FirstMem As String, SecondMem As String
    Dim count As Integer

    Do While True
        Do While True
            FirstMem = InputBox("Enter A Range under [C1:F6]")
            If Intersect(Range(FirstMem), Range("C1:F6")) Is Nothing Then
                MsgBox ("Try a valid range.")
            Else
                Exit Do
            End If
        Loop

        Range(FirstMem).Interior.Color = RGB(255, 255, 255)

        Do While True
            SecondMem = InputBox("Enter A Range under [C1:F6]")
            If Not Intersect(Range(SecondMem), Range(FirstMem)) Is Nothing Then
                MsgBox ("Try an other range this one is already taken.")
            ElseIf Intersect(Range(SecondMem), Range("C1:F6")) Is Nothing Then
                MsgBox ("Try a valid range.")
            Else
                Exit Do
            End If
        Loop

        Range(SecondMem).Interior.Color = RGB(255, 255, 255)
        If Range(SecondMem) = Range(FirstMem) Then
            Range("N1").Value = Range("N1").Value + 1
        Else
            MsgBox ("Try Again Missmatch")
            Range(FirstMem).Interior.Color = RGB(0, 0, 0)
            Range(SecondMem).Interior.Color = RGB(0, 0, 0)
        End If
        If Range("N1").Value = 12 Then
            MsgBox ("Victory, You have found all matching cards!!!")
            Exit Do
        End If
    Loop
End Sub
