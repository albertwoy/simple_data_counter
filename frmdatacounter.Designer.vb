<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmdatacounter
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.dgdc = New System.Windows.Forms.DataGridView()
        Me.btnbrws = New System.Windows.Forms.Button()
        Me.lblfile = New System.Windows.Forms.Label()
        Me.btngo = New System.Windows.Forms.Button()
        Me.pbar1 = New System.Windows.Forms.ProgressBar()
        Me.btnxcel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lbltotres = New System.Windows.Forms.Label()
        CType(Me.dgdc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgdc
        '
        Me.dgdc.AllowUserToAddRows = False
        Me.dgdc.AllowUserToDeleteRows = False
        Me.dgdc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgdc.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgdc.Location = New System.Drawing.Point(12, 113)
        Me.dgdc.Name = "dgdc"
        Me.dgdc.ReadOnly = True
        Me.dgdc.RowTemplate.Height = 24
        Me.dgdc.Size = New System.Drawing.Size(1219, 440)
        Me.dgdc.TabIndex = 0
        '
        'btnbrws
        '
        Me.btnbrws.BackColor = System.Drawing.Color.Cornsilk
        Me.btnbrws.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnbrws.Location = New System.Drawing.Point(12, 12)
        Me.btnbrws.Name = "btnbrws"
        Me.btnbrws.Size = New System.Drawing.Size(129, 35)
        Me.btnbrws.TabIndex = 1
        Me.btnbrws.Text = "Browse File ..."
        Me.btnbrws.UseVisualStyleBackColor = False
        '
        'lblfile
        '
        Me.lblfile.AutoSize = True
        Me.lblfile.Location = New System.Drawing.Point(147, 21)
        Me.lblfile.Name = "lblfile"
        Me.lblfile.Size = New System.Drawing.Size(40, 17)
        Me.lblfile.TabIndex = 2
        Me.lblfile.Text = "lblfile"
        '
        'btngo
        '
        Me.btngo.BackColor = System.Drawing.Color.PaleGreen
        Me.btngo.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btngo.Location = New System.Drawing.Point(12, 53)
        Me.btngo.Name = "btngo"
        Me.btngo.Size = New System.Drawing.Size(129, 35)
        Me.btngo.TabIndex = 3
        Me.btngo.Text = "Go >>"
        Me.btngo.UseVisualStyleBackColor = False
        '
        'pbar1
        '
        Me.pbar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbar1.Location = New System.Drawing.Point(12, 95)
        Me.pbar1.Margin = New System.Windows.Forms.Padding(4)
        Me.pbar1.Name = "pbar1"
        Me.pbar1.Size = New System.Drawing.Size(1219, 12)
        Me.pbar1.TabIndex = 94
        '
        'btnxcel
        '
        Me.btnxcel.BackColor = System.Drawing.Color.PaleGreen
        Me.btnxcel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnxcel.Location = New System.Drawing.Point(147, 53)
        Me.btnxcel.Name = "btnxcel"
        Me.btnxcel.Size = New System.Drawing.Size(177, 35)
        Me.btnxcel.TabIndex = 95
        Me.btnxcel.Text = "Export to Excel"
        Me.btnxcel.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(344, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(129, 17)
        Me.Label1.TabIndex = 96
        Me.Label1.Text = "Total Respondent :"
        '
        'lbltotres
        '
        Me.lbltotres.AutoSize = True
        Me.lbltotres.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbltotres.Location = New System.Drawing.Point(479, 62)
        Me.lbltotres.Name = "lbltotres"
        Me.lbltotres.Size = New System.Drawing.Size(65, 20)
        Me.lbltotres.TabIndex = 97
        Me.lbltotres.Text = "Label2"
        '
        'frmdatacounter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1243, 565)
        Me.Controls.Add(Me.lbltotres)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnxcel)
        Me.Controls.Add(Me.pbar1)
        Me.Controls.Add(Me.btngo)
        Me.Controls.Add(Me.lblfile)
        Me.Controls.Add(Me.btnbrws)
        Me.Controls.Add(Me.dgdc)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmdatacounter"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Counter"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.dgdc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgdc As System.Windows.Forms.DataGridView
    Friend WithEvents btnbrws As System.Windows.Forms.Button
    Friend WithEvents lblfile As System.Windows.Forms.Label
    Friend WithEvents btngo As System.Windows.Forms.Button
    Friend WithEvents pbar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnxcel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lbltotres As System.Windows.Forms.Label

End Class
