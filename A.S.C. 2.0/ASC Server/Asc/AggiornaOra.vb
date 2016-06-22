Public Class AggiornaOra

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged

        If NumericUpDown1.Value > 24 Then
            MsgBox("Il valore dell'ora dell'aggiornamento automatico non può essere maggiore di 24.", MsgBoxStyle.Critical, "Asc Server - Valore non Valido")
            NumericUpDown1.Value = 24
        End If

    End Sub

    Private Sub NumericUpDown2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown2.ValueChanged

        If NumericUpDown2.Value > 60 Then
            MsgBox("Il valore dei minuti dell'aggiornamento automatico non può essere maggiore di 60.", MsgBoxStyle.Critical, "Asc Server - Valore non Valido")
            NumericUpDown2.Value = 60
        End If

    End Sub

    Private Sub NumericUpDown3_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown3.ValueChanged

        If NumericUpDown3.Value > 60 Then
            MsgBox("Il valore dei secondi dell'aggiornamento automatico non può essere maggiore di 60.", MsgBoxStyle.Critical, "Asc Server - Valore non Valido")
            NumericUpDown3.Value = 60
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If NumericUpDown1.Value = 24 Then
                NumericUpDown1.Value = 0
            End If
            My.Settings.OraAggiornamento = NumericUpDown1.Value.ToString & ":" & NumericUpDown2.Value.ToString & ":" & NumericUpDown3.Value.ToString
            Main.MenuItem40.Text = My.Settings.OraAggiornamento
            Main.MenuItem40.Checked = True
            My.Settings.AggiornamentoAutoSelect = True
            MsgBox("Ora aggiornamento automatico impostata correttamente, è stata automaticamente selezionato l'aggiornamento automatico della data.", MsgBoxStyle.Information, "Asc Server - Ora Aggiornamento Automatico")
            Me.Close()
        Catch ex As Exception

        End Try

    End Sub
End Class