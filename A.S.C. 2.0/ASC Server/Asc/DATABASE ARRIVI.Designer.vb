﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DATABASE_ARRIVI
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DATABASE_ARRIVI))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Seleziona = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Destinatario = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MSGID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataOra = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Allegato = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GDO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Arrivi_PCN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.VisualizzaEmailToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UtilizzaMessaggioToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.SalvaEmailToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StampaEmailToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Seleziona, Me.Destinatario, Me.MSGID, Me.DataOra, Me.Allegato, Me.GDO, Me.FM, Me.Arrivi_PCN})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1028, 401)
        Me.DataGridView1.TabIndex = 0
        '
        'Seleziona
        '
        Me.Seleziona.HeaderText = "Nr."
        Me.Seleziona.Name = "Seleziona"
        Me.Seleziona.ReadOnly = True
        Me.Seleziona.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Seleziona.Width = 80
        '
        'Destinatario
        '
        Me.Destinatario.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Destinatario.HeaderText = "Destinatario(i)"
        Me.Destinatario.Name = "Destinatario"
        Me.Destinatario.ReadOnly = True
        '
        'MSGID
        '
        Me.MSGID.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.MSGID.HeaderText = "MSGID"
        Me.MSGID.Name = "MSGID"
        Me.MSGID.ReadOnly = True
        '
        'DataOra
        '
        Me.DataOra.HeaderText = "Data & Ora"
        Me.DataOra.Name = "DataOra"
        Me.DataOra.ReadOnly = True
        Me.DataOra.Width = 150
        '
        'Allegato
        '
        Me.Allegato.HeaderText = "Allegato"
        Me.Allegato.Name = "Allegato"
        Me.Allegato.ReadOnly = True
        Me.Allegato.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Allegato.Width = 80
        '
        'GDO
        '
        Me.GDO.HeaderText = "GDO"
        Me.GDO.Name = "GDO"
        Me.GDO.ReadOnly = True
        Me.GDO.Width = 120
        '
        'FM
        '
        Me.FM.HeaderText = "FM"
        Me.FM.Name = "FM"
        Me.FM.ReadOnly = True
        Me.FM.Width = 210
        '
        'Arrivi_PCN
        '
        Me.Arrivi_PCN.HeaderText = "Arrivi PCN"
        Me.Arrivi_PCN.Name = "Arrivi_PCN"
        Me.Arrivi_PCN.ReadOnly = True
        Me.Arrivi_PCN.Width = 80
        '
        'PrintDocument1
        '
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VisualizzaEmailToolStripMenuItem, Me.UtilizzaMessaggioToolStripMenuItem, Me.ToolStripMenuItem1, Me.SalvaEmailToolStripMenuItem, Me.StampaEmailToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(172, 98)
        '
        'VisualizzaEmailToolStripMenuItem
        '
        Me.VisualizzaEmailToolStripMenuItem.Name = "VisualizzaEmailToolStripMenuItem"
        Me.VisualizzaEmailToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.VisualizzaEmailToolStripMenuItem.Text = "Visualizza Email"
        '
        'UtilizzaMessaggioToolStripMenuItem
        '
        Me.UtilizzaMessaggioToolStripMenuItem.Name = "UtilizzaMessaggioToolStripMenuItem"
        Me.UtilizzaMessaggioToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.UtilizzaMessaggioToolStripMenuItem.Text = "Utilizza Messaggio"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(168, 6)
        '
        'SalvaEmailToolStripMenuItem
        '
        Me.SalvaEmailToolStripMenuItem.Name = "SalvaEmailToolStripMenuItem"
        Me.SalvaEmailToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.SalvaEmailToolStripMenuItem.Text = "Salva Email"
        '
        'StampaEmailToolStripMenuItem
        '
        Me.StampaEmailToolStripMenuItem.Name = "StampaEmailToolStripMenuItem"
        Me.StampaEmailToolStripMenuItem.Size = New System.Drawing.Size(171, 22)
        Me.StampaEmailToolStripMenuItem.Text = "Stampa Email"
        '
        'Timer1
        '
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem6})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3, Me.MenuItem4, Me.MenuItem5})
        Me.MenuItem1.Text = "File"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Salva Email"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "Stampa Email"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "-"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 3
        Me.MenuItem5.Text = "Chiudi"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 1
        Me.MenuItem6.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem7, Me.MenuItem8, Me.MenuItem9})
        Me.MenuItem6.Text = "Modifica"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 0
        Me.MenuItem7.Text = "Scegli data da visualizzare"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 1
        Me.MenuItem8.Text = "-"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 2
        Me.MenuItem9.Text = "Stampa 22-Ter"
        '
        'DATABASE_ARRIVI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1028, 401)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "DATABASE_ARRIVI"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Asc Server - Database Arrivi"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents VisualizzaEmailToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents SalvaEmailToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StampaEmailToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Seleziona As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Destinatario As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MSGID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataOra As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Allegato As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GDO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Arrivi_PCN As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UtilizzaMessaggioToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
End Class
