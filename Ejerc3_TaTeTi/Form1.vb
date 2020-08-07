Public Class Form1

    Dim jugador As String = "X"
    Dim turno As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        limpiarTablero()
    End Sub

    Private Sub limpiarTablero()
        ' dejamos un espacio para poder clickear el "caracter" en el Form
        lbl_0_0.Text = " "
        lbl_0_1.Text = " "
        lbl_0_2.Text = " "
        lbl_1_0.Text = " "
        lbl_1_1.Text = " "
        lbl_1_2.Text = " "
        lbl_2_0.Text = " "
        lbl_2_1.Text = " "
        lbl_2_2.Text = " "

    End Sub

    Private Sub lbl_0_0_Click(sender As Object, e As EventArgs) Handles lbl_0_0.Click
        jugada(sender)
    End Sub

    Private Sub lbl_0_1_Click(sender As Object, e As EventArgs) Handles lbl_0_1.Click
        jugada(sender)
    End Sub

    Private Sub lbl_0_2_Click(sender As Object, e As EventArgs) Handles lbl_0_2.Click
        jugada(sender)
    End Sub

    Private Sub lbl_1_0_Click(sender As Object, e As EventArgs) Handles lbl_1_0.Click
        jugada(sender)
    End Sub

    Private Sub lbl_1_1_Click(sender As Object, e As EventArgs) Handles lbl_1_1.Click
        jugada(sender)
    End Sub

    Private Sub lbl_1_2_Click(sender As Object, e As EventArgs) Handles lbl_1_2.Click
        jugada(sender)
    End Sub

    Private Sub lbl_2_0_Click(sender As Object, e As EventArgs) Handles lbl_2_0.Click
        jugada(sender)
    End Sub

    Private Sub lbl_2_1_Click(sender As Object, e As EventArgs) Handles lbl_2_1.Click
        jugada(sender)
    End Sub

    Private Sub lbl_2_2_Click(sender As Object, e As EventArgs) Handles lbl_2_2.Click
        jugada(sender)
    End Sub

    Private Sub jugada(ByVal sender As System.Object)
        If sender.Text.Equals(" ") Then
            sender.Text = jugador
            turno += 1
            If hayGanador() Then
                If jugador.Equals("X") Then
                    ' gana el jugador X
                    MsgBox("GANA EL JUGADOR X ", vbInformation, "Final del juego!")
                Else
                    ' gana el jugador O
                    MsgBox("GANA EL JUGADOR O ", vbInformation, "Final del juego!")

                End If
                finalJuego()

            ElseIf turno = 9 Then
                MsgBox("EMPATE!", vbEmpty, "Mensaje")
                finalJuego()
            End If

            If jugador.Equals("X") Then
                jugador = "O"
                lblTurnoJug.Text = "Turno jugador: O"
            Else
                jugador = "X"
                lblTurnoJug.Text = "Turno jugador: X"
            End If

        Else
            ' caso que sender.Text distinto a null
            MsgBox("Posición ocupada!", vbExclamation, "Mensaje")
        End If

    End Sub

    Private Function hayGanador() As Boolean
        ' determina si hay TaTeTi o no
        'Dim campoVacio As String = " "
        Dim x As Boolean

        If (lbl_0_0.Text = lbl_1_0.Text) And (lbl_1_0.Text = lbl_2_0.Text) And lbl_0_0.Text <> " " Then
            x = True
        ElseIf (lbl_0_1.Text = lbl_1_1.Text And lbl_1_1.Text = lbl_2_1.Text) And lbl_0_1.Text <> " " Then
            x = True
        ElseIf (lbl_0_2.Text = lbl_1_2.Text And lbl_1_2.Text = lbl_2_2.Text) And lbl_0_2.Text <> " " Then
            x = True
        ElseIf (lbl_0_0.Text = lbl_0_1.Text And lbl_0_1.Text = lbl_0_2.Text) And lbl_0_0.Text <> " " Then
            x = True
        ElseIf (lbl_1_0.Text = lbl_1_1.Text And lbl_1_1.Text = lbl_1_2.Text) And lbl_1_0.Text <> " " Then
            x = True
        ElseIf (lbl_2_0.Text = lbl_2_1.Text And lbl_2_1.Text = lbl_2_2.Text) And lbl_2_0.Text <> " " Then
            x = True
        ElseIf (lbl_0_0.Text = lbl_1_1.Text And lbl_1_1.Text = lbl_2_2.Text) And lbl_0_0.Text <> " " Then
            x = True
        ElseIf (lbl_0_2.Text = lbl_1_1.Text And lbl_1_1.Text = lbl_2_0.Text) And lbl_0_2.Text <> " " Then
            x = True
        Else
            x = False
        End If

        Return x

    End Function

    Private Sub finalJuego()

        If MsgBox("Deseas jugar de nuevo?", vbInformation + vbYesNo, "Mensaje") = MsgBoxResult.Yes Then
            Application.Restart() ' reinicio de aplicacion.-
        Else
            Application.Exit()
        End If

    End Sub

End Class
