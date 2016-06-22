Imports System.IO.File
Imports System.IO

Public Class UfficiAdd

    Dim email() As String
    Dim scrivi As System.IO.StreamWriter
    Dim leggi As System.IO.StreamReader
    Dim letto As String

    Public Sub AddOffice()

        Try
            leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\listauffici.txt")
            letto = leggi.ReadToEnd
            leggi.Close()
            Dim MySplitter() As String = Split(letto, vbCrLf)
            For r = 0 To UBound(MySplitter)
                If MySplitter(r).ToString = TextBox1.Text Then
                    MsgBox("Esiste un ufficio con lo stesso nome scelto, cambaire nome per l'ufficio che si desidera creare per poter proseguire.", MsgBoxStyle.Information, "Asc Server - Aggiungere Ufficio")
                    Exit Sub
                End If
            Next
            If TextBox1.Text <> "" And TextBox2.Text <> "" Then
                email = Split(TextBox2.Text, ";")
                scrivi = System.IO.File.CreateText(Application.StartupPath & "\Uffici\Ufficio" & My.Settings.Ufficio & ".txt")
                scrivi.Write(TextBox1.Text & "#")
                For r = 0 To email.Count - 1
                    If email(r) <> "" Then
                        scrivi.Write(email(r) & "#")
                    End If
                Next
                scrivi.Close()
                My.Settings.Ufficio += 1
                Uffici.TabControl1.TabPages.Clear()

                If System.IO.File.Exists(Application.StartupPath & "\Uffici\listauffici.txt") Then
                    scrivi = System.IO.File.AppendText(Application.StartupPath & "\Uffici\listauffici.txt")
                    scrivi.WriteLine(TextBox1.Text + vbCrLf)
                    scrivi.Close()
                Else
                    scrivi = System.IO.File.CreateText(Application.StartupPath & "\Uffici\listauffici.txt")
                    scrivi.WriteLine(TextBox1.Text + vbCrLf)
                    scrivi.Close()
                End If
                Me.TopMost = True
                MsgBox("L'Ufficio " & TextBox1.Text & " è stato aggiunto correttamente.", MsgBoxStyle.Information, "Asc Server - Errore Aggiunta Ufficio")
                TextBox1.Clear()
                TextBox2.Clear()
                Uffici.Timer1.Start()
                Main.LoadContacts()
                ContactsAdd.LoadUff()
                Main.SaveOthersProgressive()
            Else
                MsgBox("Mancano i dati neccessari per poter prosegurie con l'aggiunta dell'ufficio.", MsgBoxStyle.Information, "Asc Server - Errore Aggiunta Ufficio")
            End If
        Catch ex As Exception
            MsgBox("Errore: si è verificato un problema durante l'aggiunta del nuovo ufficio, ex message: " & ex.Message, MsgBoxStyle.Information, "Asc Server - Errore Aggiunta Ufficio")
        End Try

    End Sub

    Private Sub UfficiAdd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            ContactsAdd.Close()
            RimuoviUffici.Close()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        AddOffice()

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        AddOffice()

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.Close()

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            AddOffice()
        End If

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            AddOffice()
        End If

    End Sub
End Class