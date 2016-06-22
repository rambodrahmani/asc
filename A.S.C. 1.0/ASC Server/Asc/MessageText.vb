Public Class MessageText

    Dim scrivi As System.IO.StreamWriter
    Dim MysPlitter() As String
    Dim ModificheEffettuate As Integer = 0
    Dim SalvataggioinChiusura As Integer = 0
    Dim TxtBoxStartingText As String = ""

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.Close()

    End Sub

    Public Sub SaveNewData()

        Try
            MysPlitter = Split(TextBox1.Text, vbCrLf)
            TextBox1.Clear()
            For r = 0 To UBound(MysPlitter)
                If MysPlitter(r).ToString <> "Email inviata a:" Then
                    TextBox1.Text += MysPlitter(r).ToString.ToString & vbCrLf
                End If
            Next
            scrivi = System.IO.File.CreateText(My.Settings.OpenedFilePath)
            scrivi.Write(TextBox1.Text)
            scrivi.Close()
            ModificheEffettuate = 0

            If SalvataggioinChiusura = 1 Then
                Me.Close()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        Try
            My.Settings.x = 7
            Accesso.Show()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MessageText_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Try
            If TxtBoxStartingText <> TextBox1.Text Then
                If ModificheEffettuate = 1 Then
                    Dim AlertMessage = MsgBox("Il testo della email è stato modificato, si desidera salvare?", MsgBoxStyle.YesNo, "Asc - Salvataggio Modifiche")
                    If AlertMessage = 6 Then
                        e.Cancel = True
                        SalvataggioinChiusura = 1
                        My.Settings.x = 7
                        Accesso.Show()
                    End If
                Else
                    My.Settings.OpenedFilePath = ""
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub MessageText_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TextBox1.SelectionStart = 0
        TextBox1.SelectionLength = 0
        TxtBoxStartingText = TextBox1.Text

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        ModificheEffettuate = 1

    End Sub

End Class