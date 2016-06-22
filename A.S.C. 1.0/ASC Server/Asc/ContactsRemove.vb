Imports System.IO.File

Public Class ContactsRemove

    Dim contattiufficio() As String
    Dim infoufficio() As String
    Dim leggi As System.IO.StreamReader
    Dim scrivi As System.IO.StreamWriter
    Dim letto As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            ListBox2.Items.Remove(ListBox2.SelectedItem)
            RemoveContct()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ContactsRemove_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadList()

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

        LoadSecList()

    End Sub

    Public Sub LoadList()

        Try
            ListBox1.Items.Clear()
            For r = 0 To My.Settings.Ufficio - 1
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                infoufficio = Split(letto, "#")
                ListBox1.Items.Add(infoufficio(0))
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub LoadSecList()

        Try
            ListBox2.Items.Clear()
            For r = 0 To My.Settings.Ufficio - 1
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                contattiufficio = Split(letto, "#")
                If contattiufficio(0) = ListBox1.SelectedItem.ToString Then
                    For z = 1 To UBound(contattiufficio) - 1
                        ListBox2.Items.Add(contattiufficio(z))
                    Next
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub RemoveContct()

        Try
            For r = 0 To My.Settings.Ufficio - 1
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                infoufficio = Split(letto, "#")
                If infoufficio(0) = ListBox1.SelectedItem.ToString Then
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                    scrivi.Write(ListBox1.SelectedItem.ToString & "#")
                    scrivi.Close()
                    For z = 0 To ListBox2.Items.Count - 1
                        ListBox2.SelectedIndex = z
                        scrivi = System.IO.File.AppendText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                        scrivi.Write(ListBox2.SelectedItem.ToString & "#")
                        scrivi.Close()
                    Next
                End If
            Next

            Uffici.Timer1.Start()
        Catch ex As Exception

        End Try
        
    End Sub

End Class