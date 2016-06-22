Public Class ExchangeSettings

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If NumericUpDown1.Value <> 0 And NumericUpDown2.Value <> 0 Then
                If NumericUpDown1.Value < 10 Then
                    Dim AlertMessage = MsgBox("Per il controllo della presenza di code è stato impostato un valore inferiore a 10 secondi. Non è consigliato diminuire troppo questo valore e si consiglia di impostare un valore pari almeno a 10 secondi.", MsgBoxStyle.YesNo, "Asc Server - Impostazioni Exchange")
                    If AlertMessage = 6 Then
                        My.Settings.SettingsFrequenzaControllo = NumericUpDown1.Value
                        Main.client.Send(2, "NewSettingsFromServerForExchange" & "|" & My.Settings.SettingsFrequenzaControllo & "|" & My.Settings.SettingsFrequenzaInvio & "|")
                        MsgBox("Impostazioni salvate correttamente.", MsgBoxStyle.Information, "Asc Server - Impostazioni Salvate")
                        Me.Close()
                    End If
                Else
                    My.Settings.SettingsFrequenzaControllo = NumericUpDown1.Value
                    My.Settings.SettingsFrequenzaInvio = NumericUpDown2.Value
                    Main.client.Send(2, "NewSettingsFromServerForExchange" & "|" & My.Settings.SettingsFrequenzaControllo & "|" & My.Settings.SettingsFrequenzaInvio & "|")
                    MsgBox("Impostazioni salvate correttamente.", MsgBoxStyle.Information, "Asc Server - Impostazioni Salvate")
                    Me.Close()
                End If
            Else
                MsgBox("Uno dei due valori è stato impostato su 0, non è possibile impostare alcun valore su 0. Impossibile salvare queste impostazioni.", MsgBoxStyle.Critical, "Asc Server - Errore Importazioni")
            End If
        Catch ex As Exception

        End Try

    End Sub

End Class