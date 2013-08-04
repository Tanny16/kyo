Attribute VB_Name = "Module1"
Sub main()
Dim zhs(10) As String
Dim zhn(10) As Single
Dim temp_s, temp_s2 As Single
Dim temp_l As String
Dim x, y, z, czfx As Integer
Dim line_input, czf As String
Open "input.txt" For Input As #1
Do
    Line Input #1, line_input
    If Mid$(line_input, 1, 1) <> "1" Then Exit Do
    zhs(x) = Mid$(line_input, 2, InStr(line_input, " = ") - 2)
    zhn(x) = CSng(Mid$(line_input, InStr(line_input, " = ") + 3, Len(line_input) - InStr(line_input, " = ") - 4))
    x = x + 1
Loop

x = 0
Open "output.txt" For Output As #2
Print #2, "595915575@qq.com"
Print #2,
Do
    If Not (EOF(1)) Then
        Line Input #1, line_input
        line_input = line_input & " "
    Else
        Exit Do
    End If
    x = 0: y = 0: z = 1: czfx = 0
    Do
        y = InStr(z, line_input, " ")
        temp_s = CSng(Mid$(line_input, z, y - z))
        z = InStr(y + 1, line_input, " ")
        temp_l = Mid$(line_input, y, z - y)
        If temp_l = " feet" Then temp_l = " foot"
        y = z
        x = 0
        Do
            If InStr(temp_l, zhs(x)) = 1 Then
                If czfx = 0 Then
                    temp_s2 = temp_s * zhn(x)
                ElseIf czfx = 1 Then
                    temp_s2 = temp_s2 + temp_s * zhn(x)
                ElseIf czfx = 2 Then
                    temp_s2 = temp_s2 - temp_s * zhn(x)
                End If
                Exit Do
            End If
            x = x + 1
        Loop
  
        If y = Len(line_input) Then
            czfx = 0
            Print #2, Format$("" & temp_s2, "0.00") & " m"
            temp_s2 = 0
            Exit Do
        Else
            czf = Mid$(line_input, y + 1, InStr(y + 1, line_input, " ") - y - 1)
            If czf = "+" Then
                czfx = 1
            ElseIf czf = "-" Then
                czfx = 2
            End If
            z = z + 3
        End If
    Loop
Loop
Close #2
Close #1
End Sub
