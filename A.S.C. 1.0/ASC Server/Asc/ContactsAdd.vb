Imports System.IO.File

Public Class ContactsAdd

    Dim items() As String
    Dim infoufficio() As String
    Dim scrivi As System.IO.StreamWriter
    Dim leggi As System.IO.StreamReader
    Dim letto As String
    Dim nomeufficio As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        nomeufficio = ListBox1.SelectedItem.ToString
        If nomeufficio <> "" Then
            Try
                For r = 0 To My.Settings.Ufficio - 1
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    infoufficio = Split(letto, "#")
                    If infoufficio(0) = nomeufficio Then
                        scrivi = System.IO.File.AppendText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                        scrivi.WriteLine(TextBox1.Text & "#")
                        scrivi.Close()
                    End If
                Next
                MsgBox("Contatto aggiunto all'ufficio " & nomeufficio & ".", MsgBoxStyle.Information)
                Uffici.Timer1.Start()
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub ContactsAdd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            UfficiAdd.Close()
            RimuoviUffici.Close()
            LoadUff()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadUff()

        Try
            ListBox1.Items.Clear()
            leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\listauffici.txt")
            letto = leggi.ReadToEnd
            leggi.Close()
            items = Split(letto, vbCrLf)
            For r = 0 To UBound(items) - 1
                If items(r) <> "" Then
                    ListBox1.Items.Add(items(r))
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Button1.PerformClick()
        End If

    End Sub

End Class