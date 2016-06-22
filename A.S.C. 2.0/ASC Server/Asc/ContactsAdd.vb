Imports System.IO.File

Public Class ContactsAdd

    Dim items() As String
    Dim infoufficio() As String
    Dim scrivi As System.IO.StreamWriter
    Dim leggi As System.IO.StreamReader
    Dim letto As String
    Dim nomeufficio As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If ListBox1.SelectedItems.Count = 1 Then
            nomeufficio = ListBox1.SelectedItem.ToString
        Else
            MsgBox("Impossibile aggiungere il contatto, devi selezionare l'ufficio in cui aggiungerlo.", MsgBoxStyle.Critical, "Asc Server - Errore Aggiunta Contatto")
        End If
        If nomeufficio <> "" Then
            Try
                For r = 0 To My.Settings.Ufficio - 1
                    leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                    letto = leggi.ReadToEnd
                    leggi.Close()
                    infoufficio = Split(letto, "#")
                    If infoufficio(0) = nomeufficio Then
                        Dim MySplitter() As String = Split(letto, "#")
                        scrivi = System.IO.File.CreateText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                        For x = 0 To UBound(MySplitter)
                            If MySplitter(x).ToString <> "" Then
                                scrivi.Write(MySplitter(x).ToString & "#")
                            Else
                                If x = MySplitter.Count - 1 Then
                                    scrivi.Write(TextBox1.Text & "#")
                                End If
                            End If
                        Next
                        scrivi.Close()
                    End If
                Next
                MsgBox("Contatto aggiunto all'ufficio " & nomeufficio & ".", MsgBoxStyle.Information, "Asc Server - Contatto Aggiunto Correttamente")
                Uffici.Timer1.Start()
            Catch ex As Exception
                MsgBox(ex.Message)
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