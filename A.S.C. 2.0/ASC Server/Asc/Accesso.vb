Public Class Accesso

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If CheckBox1.Checked = True Then
                My.Settings.ID = TextBox1.Text
            End If
            If CheckBox2.Checked = True Then
                My.Settings.PASS = TextBox2.Text
            End If

            If TextBox1.Text = My.Settings.NomeUtente Then
                If TextBox2.Text = My.Settings.PassUtente Then
                    If My.Settings.x = 1 Then
                        UfficiAdd.Show()
                        My.Settings.x = 0
                        Me.Close()
                    ElseIf My.Settings.x = 2 Then
                        RimuoviUffici.Show()
                        My.Settings.x = 2
                        Me.Close()
                    ElseIf My.Settings.x = 3 Then
                        ContactsAdd.Show()
                        My.Settings.x = 0
                        Me.Close()
                    ElseIf My.Settings.x = 4 Then
                        ContactsRemove.Show()
                        My.Settings.x = 0
                        Me.Close()
                    ElseIf My.Settings.x = 5 Then
                        My.Settings.x = 0
                        Main.RequestSettings()
                        Me.Close()
                    ElseIf My.Settings.x = 6 Then
                        If My.Settings.DTBArrivi = 1 Then
                            My.Settings.DTBArrivi = 0
                            DATABASE_ARRIVI.DeleteSelectedEmail()
                        End If
                        If My.Settings.DTBPartenze = 1 Then
                            My.Settings.DTBPartenze = 0
                            DATABASE_PARTENZE.DeleteSelectedEmail()
                        End If
                        My.Settings.x = 0
                        Me.Close()
                    ElseIf My.Settings.x = 7 Then
                        MessageText.SaveNewData()
                        My.Settings.x = 0
                        Me.Close()
                    ElseIf My.Settings.x = 8 Then
                        My.Settings.x = 0
                        Utente.Show()
                        Me.Close()
                    End If
                Else
                    MsgBox("Password Utente Errata", MsgBoxStyle.Information, "Asc Server - Acesso Uffici")
                End If
            Else
                MsgBox("Nome Utente Errato", MsgBoxStyle.Information, "Asc Server - Acesso Uffici")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Accesso_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TextBox1.Text = My.Settings.ID
        TextBox2.Text = My.Settings.PASS
        If My.Settings.ID <> "" Then
            CheckBox1.Checked = True
        End If
        If My.Settings.PASS <> "" Then
            CheckBox2.Checked = True
        End If

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If

    End Sub

End Class