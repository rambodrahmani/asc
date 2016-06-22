Public Class Settings

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If NumericUpDown1.Value <> 0 And NumericUpDown2.Value <> 0 Then
                If NumericUpDown1.Value < 10 Then
                    Dim AlertMessage = MsgBox("Per il controllo della presenza di code è stato impostato un valore inferiore a 10 secondi. Non è consigliato diminuire troppo questo valore e si consiglia di impostare un valore pari almeno a 10 secondi.", MsgBoxStyle.YesNo, "Asc Exchange - Impostazioni")
                    If AlertMessage = 6 Then
                        My.Settings.FrequenzaControllo = (NumericUpDown1.Value * 1000)
                        My.Settings.FrequenzaInvio = (NumericUpDown2.Value * 1000)
                        MsgBox("Impostazioni salvate correttamente.", MsgBoxStyle.Information, "Asc Exchange - Impostazioni Salvate")
                        Form1.SetSettings()
                        Me.Close()
                    End If
                Else
                    My.Settings.FrequenzaControllo = (NumericUpDown1.Value * 1000)
                    My.Settings.FrequenzaInvio = (NumericUpDown2.Value * 1000)
                    MsgBox("Impostazioni salvate correttamente.", MsgBoxStyle.Information, "Asc Exchange - Impostazioni Salvate")
                    Form1.SetSettings()
                    Me.Close()
                End If
            Else
                MsgBox("Uno dei due valori è stato impostato su 0, non è possibile impostare alcun valore su 0. Impossibile salvare queste impostazioni.", MsgBoxStyle.Critical, "Asc Exchange - Errore Importazioni")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Settings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            NumericUpDown1.Value = (My.Settings.FrequenzaControllo / 1000)
            NumericUpDown2.Value = (My.Settings.FrequenzaInvio / 1000)
        Catch ex As Exception

        End Try

    End Sub
End Class