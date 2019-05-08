'imports
Imports System.IO
Imports System.Net

Public Class Form1

    'variables
    Public Shared obj As PictureBox
    Public Shared table(5, 6) As Integer
    Public Shared objName, column, row As String
    Public Shared player As Integer

    'pictube box clicks
    Public Sub ClickButton(sender As Object, e As EventArgs) Handles c50.Click, c51.Click,
    c52.Click, c53.Click, c54.Click, c55.Click, c56.Click, c40.Click, c41.Click, c42.Click,
    c43.Click, c44.Click, c45.Click, c46.Click, c30.Click, c31.Click, c32.Click, c33.Click,
    c34.Click, c35.Click, c36.Click, c20.Click, c21.Click, c22.Click, c23.Click, c24.Click,
    c25.Click, c26.Click, c10.Click, c11.Click, c12.Click, c13.Click, c14.Click, c15.Click,
    c16.Click, c00.Click, c01.Click, c02.Click, c03.Click, c04.Click, c05.Click, c06.Click

        obj = CType(sender, PictureBox)
        objName = obj.Name

        column = objName.Chars(1)
        row = objName.Chars(2)
        lblPos.Text = "Pos: " & column & row

        If table(column, row) <> 0 Then
        Else
            table(column, row) = player
            If ChangeColor(obj, player) Then
                If player = 1 Then
                    player = 2
                ElseIf player = 2 Then
                    player = 1
                Else
                    'jugador 3 )?? XD
                End If
            End If
        End If

        lblPlayer.Text = "Current Player: " & ConvPlayer(player)

    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'set variable data
        player = 1
        lblPlayer.Text = "Current Player: " & ConvPlayer(player)

        For i As Integer = 0 To 5
            For o As Integer = 0 To 6
                table(i, o) = 0
            Next
        Next
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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

End Class

Module methods

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
        End Select
        Return False
    End Function

End Module