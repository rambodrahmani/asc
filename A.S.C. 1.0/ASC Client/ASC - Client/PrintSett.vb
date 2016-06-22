Public Class PrintSett

    Private Sub PrintSett_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If My.Settings.SavePrintSett = 1 Then
            CheckBox1.Checked = True
        End If

        If My.Settings.Top <> "0" Then
            TextBox1.Text = My.Settings.Top
        Else
            TextBox1.Text = 80
        End If

        If My.Settings.Left <> "0" Then
            TextBox2.Text = My.Settings.Left
        Else
            TextBox2.Text = 90
        End If

        If My.Settings.Width <> "0" Then
            TextBox3.Text = My.Settings.Width
        Else
            TextBox3.Text = 2100
        End If

        If My.Settings.Height <> "0" Then
            TextBox4.Text = My.Settings.Height
        Else
            TextBox4.Text = 2780
        End If

        If My.Settings.NumRighePerPagina <> "0" Then
            NumericUpDown1.Value = My.Settings.NumRighePerPagina
        Else
            NumericUpDown1.Value = 50
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If CheckBox1.Checked = True Then
            My.Settings.Top = TextBox1.Text
            My.Settings.Left = TextBox2.Text
            My.Settings.Width = TextBox3.Text
            My.Settings.Height = TextBox4.Text
            My.Settings.NumRighePerPagina = NumericUpDown1.Value
            My.Settings.SavePrintSett = 1
            MsgBox("Preferenze stampa memorizzate correttemente, da quetso momento in poi usa il MainMenu alla voce 'Modifica' per modificare le tue preferenze.", MsgBoxStyle.Information, "Asc - Preferenze Stampa")
        Else
            My.Settings.Top = TextBox1.Text
            My.Settings.Left = TextBox2.Text
            My.Settings.Width = TextBox3.Text
            My.Settings.Height = TextBox4.Text
            My.Settings.NumRighePerPagina = NumericUpDown1.Value
            My.Settings.SavePrintSett = 0
        End If

        If My.Settings.x = 9 Then
            Form1.Stampa()
            My.Settings.x = 0
        ElseIf My.Settings.x = 10 Then
            Form1.Stampa2()
            My.Settings.x = 0
        End If

        Me.Close()

    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged

        My.Settings.NumRighePerPagina = NumericUpDown1.Value

    End Sub

End Class