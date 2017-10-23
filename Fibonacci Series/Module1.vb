Module Module1

    Sub Main()

        Dim Sum As Integer = 0
        Dim A As Integer = 0
        Dim B As Integer = 1

        While True

            A = A + B

            B = A - B

            If A >= 4000000 Then
                Exit While
            Else
                If (A Mod 2 = 0) Then

                    Sum += A

                End If
            End If

        End While

        MsgBox(" The sum of the even-valued terms is " & Sum & "")

    End Sub


   

End Module
