'imports
Imports System.IO
Imports System.Net

Public Class Form1

    'variables
    Public obj As PictureBox
    Public table(5, 6) As Integer
    Public column, row, namae As String
    Public player As Integer

    'pictube box clicks
    Public Sub ClickButton(sender As Object, e As EventArgs) Handles c50.Click, c51.Click,
    c52.Click, c53.Click, c54.Click, c55.Click, c56.Click, c40.Click, c41.Click, c42.Click,
    c43.Click, c44.Click, c45.Click, c46.Click, c30.Click, c31.Click, c32.Click, c33.Click,
    c34.Click, c35.Click, c36.Click, c20.Click, c21.Click, c22.Click, c23.Click, c24.Click,
    c25.Click, c26.Click, c10.Click, c11.Click, c12.Click, c13.Click, c14.Click, c15.Click,
    c16.Click, c00.Click, c01.Click, c02.Click, c03.Click, c04.Click, c05.Click, c06.Click

        obj = CType(sender, PictureBox)
        column = obj.Name.Chars(1)
        row = obj.Name.Chars(2)
        lblPos.Text = "Pos: " & column & row

        For i As Integer = 5 To 0 Step -1
            If table(i, row) = 0 Then
                namae = "c" & i & row
                obj = Controls(namae)
                If ChangeColor(obj, player) Then
                    table(i, row) = player
                    If player = 1 Then
                        player = 2
                    ElseIf player = 2 Then
                        player = 1
                    Else
                        'jugador 3 )?? XD
                    End If
                    CheckWin()
                    UpdateLB()
                End If
                Exit For
            End If
        Next

        lblPlayer.Text = "Current Player: " & player & " - " & ConvPlayer(player)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'set variable data
        player = 1
        lblPlayer.Text = "Current Player: " & player & " - " & ConvPlayer(player)

        ResetTable()
        UpdateLB()
    End Sub

    Public Sub UpdateLB()
        'write 2d array to listbox // debug
        Dim output As String
        ListBox1.Items.Clear()
        For i = 0 To 5
            output = ""
            For k As Integer = 0 To 6
                output &= table(i, k) & " "
            Next k
            ListBox1.Items.Add(output)
        Next i
    End Sub

    Public Sub ResetTable()
        For i As Integer = 0 To 5
            For o As Integer = 0 To 6
                table(i, o) = 0
            Next
        Next
    End Sub

    Public Function ConvPlayer(a As Integer)
        If a = 1 Then
            Return "Red"
        ElseIf a = 2 Then
            Return "Blue"
        End If
        Return "weno cagamos wachin"
    End Function

    Public Sub WriteTxt()
        Using writer As New StreamWriter("C:\line4\data.txt", True)
            writer.WriteLine()
        End Using
    End Sub

    Public Function ChangeColor(b As PictureBox, i As Integer)

        Select Case i
            Case 1
                b.Image = My.Resources.b
                Return True
            Case 2
                b.Image = My.Resources.c
                Return True
            Case 0
                b.Image = My.Resources.a
                Return True
        End Select
        Return False

    End Function


    Public Sub CheckWin()

        Dim win1 As Integer = 0
        Dim win2 As Integer = 0
        Dim c As Integer = 0

#Region "Horizontal"
        For i As Integer = 5 To 0 Step -1
            win1 = 0
            win2 = 0
            For o As Integer = 0 To 6 Step 1
                If table(i, o) = 1 Then
                    win1 += 1
                    win2 = 0
                ElseIf table(i, o) = 2 Then
                    win2 += 1
                    win1 = 0
                ElseIf table(i, o) = 0 Then
                    win1 = 0
                    win2 = 0
                End If
                If win1 = 4 Then
                    Win("player 1 horizontal")
                    Exit Sub
                ElseIf win2 = 4 Then
                    Win("player 2 horizontal")
                    Exit Sub
                End If
            Next
        Next
        c = 0
#End Region
#Region "Vertical"
        For o As Integer = 0 To 6 Step 1
            win1 = 0
            win2 = 0
            For i As Integer = 5 To 0 Step -1
                If table(i, o) = 1 Then
                    win1 += 1
                    win2 = 0
                ElseIf table(i, o) = 2 Then
                    win2 += 1
                    win1 = 0
                ElseIf table(i, o) = 0 Then
                    win1 = 0
                    win2 = 0
                End If
                If win1 = 4 Then
                    Win("player 1 vertical")
                    Exit Sub
                ElseIf win2 = 4 Then
                    Win("player 2 vertical")
                    Exit Sub
                End If
            Next
        Next
        win1 = 0
        win2 = 0
#End Region
#Region "Diagonal left-right"
        For h As Integer = 0 To 5
            If table(h, c) = 1 Then
                win1 += 1
                win2 = 0
            ElseIf table(h, c) = 2 Then
                win2 += 1
                win1 = 0
            ElseIf table(h, c) = 0 Then
                win1 = 0
                win2 = 0
            End If
            If win1 = 4 Then
                Win("player 1 diagonal")
                Exit Sub
            ElseIf win2 = 4 Then
                Win("player 2 diagonal")
                Exit Sub
            End If
            c += 1
        Next
        win1 = 0
        win2 = 0
        c = 0
        For h As Integer = 1 To 5
            If table(h, c) = 1 Then
                win1 += 1
                win2 = 0
            ElseIf table(h, c) = 2 Then
                win2 += 1
                win1 = 0
            ElseIf table(h, c) = 0 Then
                win1 = 0
                win2 = 0
            End If
            If win1 = 4 Then
                Win("player 1 diagonal")
                Exit Sub
            ElseIf win2 = 4 Then
                Win("player 2 diagonal")
                Exit Sub
            End If
            c += 1
        Next
        win1 = 0
        win2 = 0
        c = 0
        For h As Integer = 2 To 5
            If table(h, c) = 1 Then
                win1 += 1
                win2 = 0
            ElseIf table(h, c) = 2 Then
                win2 += 1
                win1 = 0
            ElseIf table(h, c) = 0 Then
                win1 = 0
                win2 = 0
            End If
            If win1 = 4 Then
                Win("player 1 diagonal")
                Exit Sub
            ElseIf win2 = 4 Then
                Win("player 2 diagonal")
                Exit Sub
            End If
            c += 1
        Next
        win1 = 0
        win2 = 0
        c = 1
        For h As Integer = 0 To 5
            If table(h, c) = 1 Then
                win1 += 1
                win2 = 0
            ElseIf table(h, c) = 2 Then
                win2 += 1
                win1 = 0
            ElseIf table(h, c) = 0 Then
                win1 = 0
                win2 = 0
            End If
            If win1 = 4 Then
                Win("player 1 diagonal")
                Exit Sub
            ElseIf win2 = 4 Then
                Win("player 2 diagonal")
                Exit Sub
            End If
            c += 1
        Next
        win1 = 0
        win2 = 0
        c = 2
        For h As Integer = 0 To 4
            If table(h, c) = 1 Then
                win1 += 1
                win2 = 0
            ElseIf table(h, c) = 2 Then
                win2 += 1
                win1 = 0
            ElseIf table(h, c) = 0 Then
                win1 = 0
                win2 = 0
            End If
            If win1 = 4 Then
                Win("player 1 diagonal")
                Exit Sub
            ElseIf win2 = 4 Then
                Win("player 2 diagonal")
                Exit Sub
            End If
            c += 1
        Next
        win1 = 0
        win2 = 0
        c = 3
        For h As Integer = 0 To 3
            If table(h, c) = 1 Then
                win1 += 1
                win2 = 0
            ElseIf table(h, c) = 2 Then
                win2 += 1
                win1 = 0
            ElseIf table(h, c) = 0 Then
                win1 = 0
                win2 = 0
            End If
            If win1 = 4 Then
                Win("player 1 diagonal")
                Exit Sub
            ElseIf win2 = 4 Then
                Win("player 2 diagonal")
                Exit Sub
            End If
            c += 1
        Next
        win1 = 0
        win2 = 0
        c = 0


#End Region
#Region "Diagonal right-left"



#End Region

    End Sub

    Public Sub Win(a As String)

        MsgBox(a)
        ResetAll()

    End Sub
    Public Sub ResetAll()

        Dim change As String
        Dim change2 As Integer

        ResetTable()
        UpdateLB()

        For i As Integer = 0 To 56
            If i > 6 And i < 10 Then
                Continue For
            ElseIf i >= 10 Then
                change = i
                change = change.Chars(1)
                change2 = change
                If change2 > 6 Then
                    Continue For
                End If
            End If

            If i < 10 Then
                obj = Controls("c0" & i)
                ChangeColor(obj, 0)
            Else
                obj = Controls("c" & i)
                ChangeColor(obj, 0)
            End If
        Next

    End Sub

End Class
