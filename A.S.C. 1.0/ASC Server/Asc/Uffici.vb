Imports System.IO.File
Imports System.IO.Directory
Imports System.IO

Public Class Uffici

    Dim change As Integer = 0
    Dim numLista As Integer
    Dim daCercare As String
    Dim infoufficio() As String
    Dim letto As String
    Dim leggi As System.IO.StreamReader
    Dim scrivi As System.IO.StreamWriter
    Private tabPage1 As TabPage
    Private list As ListBox
    Dim x As Integer

    Private Sub Uffici_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        My.Settings.UfficiIsActive = 0

    End Sub

    Private Sub Uffici_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        My.Settings.UfficiIsActive = 1
        My.Settings.x = 0
        Timer1.Start()

    End Sub

    Public Sub Rimozione()

        Try
            System.IO.File.Delete(Application.StartupPath & "\Uffici\listauffici.txt")

            For r = 0 To My.Settings.Ufficio - 1
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                letto = leggi.ReadToEnd()
                leggi.Close()
                infoufficio = Split(letto, "#")
                If infoufficio(0) = My.Settings.SelectedOffice Then
                    System.IO.File.Delete(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                End If
            Next

            My.Settings.SelectedOffice = ""
            My.Settings.Ufficio -= 1

            For Each oldfilename As String In My.Computer.FileSystem.GetFiles(Application.StartupPath & "\Uffici")
                My.Computer.FileSystem.RenameFile(oldfilename, "ciao" & change & ".txt")
                change += 1
            Next

            change = 0

            For Each oldfilename As String In My.Computer.FileSystem.GetFiles(Application.StartupPath & "\Uffici")
                My.Computer.FileSystem.RenameFile(oldfilename, "Ufficio" & change & ".txt")
                change += 1
            Next

            Timer1.Start()

            Main.SaveOthersProgressive()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub creaLista()

        Try
            scrivi = System.IO.File.CreateText(Application.StartupPath & "\Uffici\listauffici.txt")

            TabControl1.SelectedIndex = 0

            For r = 0 To My.Settings.Ufficio - 1
                TabControl1.SelectedIndex = r
                scrivi.Write(TabControl1.SelectedTab.Text + vbCrLf)
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                infoufficio = Split(letto, "#")
                If infoufficio.Count > 2 Then
                    TabControl1.SelectedTab.Text += " +"
                End If
            Next
            scrivi.Close()

            Main.LoadContacts()

            TabControl1.SelectedIndex = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TabControl1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TabControl1.MouseDown

        Try
            Dim punto As Point = MousePosition
            If e.Button = Windows.Forms.MouseButtons.Right Then
                ContextMenuStrip1.Show(punto)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub AggiungiUfficioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AggiungiUfficioToolStripMenuItem.Click

        Accesso.Show()
        My.Settings.x = 1

    End Sub

    Private Sub RimuoviUfficioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RimuoviUfficioToolStripMenuItem.Click

        My.Settings.x = 2
        Accesso.Show()

    End Sub

    Private Sub AggiungiContattoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AggiungiContattoToolStripMenuItem.Click

        My.Settings.x = 3
        Accesso.Show()

    End Sub

    Private Sub RimuoviContattoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RimuoviContattoToolStripMenuItem.Click

        My.Settings.x = 4
        Accesso.Show()

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        My.Settings.x = 1
        Accesso.Show()

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        My.Settings.x = 2
        Accesso.Show()

    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click

        My.Settings.x = 3
        Accesso.Show()

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        My.Settings.x = 4
        Accesso.Show()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try
            TabControl1.TabPages.Clear()

            For r = 0 To My.Settings.Ufficio - 1
                leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\Ufficio" & r & ".txt")
                letto = leggi.ReadToEnd
                leggi.Close()
                infoufficio = Split(letto, "#")
                tabPage1 = New TabPage
                list = New ListBox
                tabPage1.Text = infoufficio(0)
                TabControl1.TabPages.Add(tabPage1)
                list.Dock = DockStyle.Fill
                tabPage1.Controls.Add(list)
                For z = 1 To infoufficio.Count - 1
                    If infoufficio(z) <> "" Then
                        list.Items.Add(infoufficio(z))
                    End If
                Next
            Next

            Timer1.Stop()
            creaLista()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click

        Try
            If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim DirectoryName As New DirectoryInfo(FolderBrowserDialog1.SelectedPath)
                If DirectoryName.Name.ToString = "Uffici" Then
                    Dim FilesCollection() As String = System.IO.Directory.GetFiles(FolderBrowserDialog1.SelectedPath)
                    My.Settings.Ufficio = FilesCollection.Count - 1
                    System.IO.Directory.Delete(Application.StartupPath & "\Uffici", True)
                    System.IO.Directory.Move(FolderBrowserDialog1.SelectedPath, Application.StartupPath & "\Uffici")
                    MsgBox("Importazione eseguita con successo.", MsgBoxStyle.Information, "Asc - Importa Uffici")
                    Timer1.Start()
                Else
                    MsgBox("Per poter importare gli uffici bisogna selezionare la cartella" & """" & "Uffici" & """" & ", contenente la lista degli uffici.", MsgBoxStyle.Information, "Asc - Importa Uffici")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox("Errore durante l'importazione, ex message: " & ex.Message, MsgBoxStyle.Information, "Asc - Importa Uffici")
        End Try

    End Sub
End Class