Public Class RimuoviUffici

    Dim leggi As System.IO.StreamReader
    Dim letto As String = ""
    Dim ufficio() As String

    Public Sub CaricaUffici()

        Try
            leggi = System.IO.File.OpenText(Application.StartupPath & "\Uffici\listauffici.txt")
            letto = leggi.ReadToEnd
            leggi.Close()
            ufficio = Split(letto, vbCrLf)
            For r = 0 To UBound(ufficio)
                ListView1.Items.Add(ufficio(r).ToString)
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RimuoviUffici_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        CaricaUffici()

        Try
            UfficiAdd.Close()
            ContactsAdd.Close()
        Catch ex As Exception

        End Try

    End Sub

    Public Sub RimuoviUfficio()

        If My.Settings.UfficiIsActive = 1 Then
            Dim listview1selecteditem As ListViewItem
            For Each listview1selecteditem In ListView1.SelectedItems
                My.Settings.SelectedOffice = listview1selecteditem.Text.ToString
                listview1selecteditem.Remove()
            Next
            Me.TopMost = True
            Uffici.Rimozione()
        Else
            MsgBox("Aprire il form degli uffici per poter proseguire. Rimozione annullata.", MsgBoxStyle.Information, "Asc - Rimozione Ufficio")
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        RimuoviUfficio()

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        RimuoviUfficio()

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.Close()

    End Sub
End Class