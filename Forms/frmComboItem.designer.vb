<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmComboItem
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
        Me.txtField06 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtField05 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtField03 = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cmdButton03 = New System.Windows.Forms.Button()
        Me.cmdButton02 = New System.Windows.Forms.Button()
        Me.cmdButton01 = New System.Windows.Forms.Button()
        Me.cmdButton00 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.optButton02 = New System.Windows.Forms.RadioButton()
        Me.optButton01 = New System.Windows.Forms.RadioButton()
        Me.optButton00 = New System.Windows.Forms.RadioButton()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtField06
        '
        Me.txtField06.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField06.Location = New System.Drawing.Point(80, 29)
        Me.txtField06.Name = "txtField06"
        Me.txtField06.Size = New System.Drawing.Size(272, 20)
        Me.txtField06.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 33)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(55, 13)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Item Desc"
        '
        'txtField05
        '
        Me.txtField05.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField05.Location = New System.Drawing.Point(80, 7)
        Me.txtField05.Name = "txtField05"
        Me.txtField05.Size = New System.Drawing.Size(136, 20)
        Me.txtField05.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 11)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(47, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Barcode"
        '
        'txtField03
        '
        Me.txtField03.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField03.Location = New System.Drawing.Point(80, 51)
        Me.txtField03.Name = "txtField03"
        Me.txtField03.Size = New System.Drawing.Size(88, 20)
        Me.txtField03.TabIndex = 5
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 55)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(43, 13)
        Me.Label14.TabIndex = 4
        Me.Label14.Text = "Min Qty"
        '
        'cmdButton03
        '
        Me.cmdButton03.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdButton03.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdButton03.Location = New System.Drawing.Point(320, 1)
        Me.cmdButton03.Name = "cmdButton03"
        Me.cmdButton03.Size = New System.Drawing.Size(53, 53)
        Me.cmdButton03.TabIndex = 4
        Me.cmdButton03.Text = "Close"
        Me.cmdButton03.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdButton03.UseVisualStyleBackColor = True
        '
        'cmdButton02
        '
        Me.cmdButton02.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdButton02.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdButton02.Location = New System.Drawing.Point(267, 1)
        Me.cmdButton02.Name = "cmdButton02"
        Me.cmdButton02.Size = New System.Drawing.Size(53, 53)
        Me.cmdButton02.TabIndex = 2
        Me.cmdButton02.Text = "Search"
        Me.cmdButton02.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdButton02.UseVisualStyleBackColor = True
        '
        'cmdButton01
        '
        Me.cmdButton01.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdButton01.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdButton01.Location = New System.Drawing.Point(214, 1)
        Me.cmdButton01.Name = "cmdButton01"
        Me.cmdButton01.Size = New System.Drawing.Size(53, 53)
        Me.cmdButton01.TabIndex = 1
        Me.cmdButton01.Text = "Save"
        Me.cmdButton01.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdButton01.UseVisualStyleBackColor = True
        '
        'cmdButton00
        '
        Me.cmdButton00.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdButton00.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdButton00.Location = New System.Drawing.Point(161, 1)
        Me.cmdButton00.Name = "cmdButton00"
        Me.cmdButton00.Size = New System.Drawing.Size(53, 53)
        Me.cmdButton00.TabIndex = 0
        Me.cmdButton00.Text = "Add"
        Me.cmdButton00.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdButton00.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdButton03)
        Me.Panel1.Controls.Add(Me.cmdButton02)
        Me.Panel1.Controls.Add(Me.cmdButton00)
        Me.Panel1.Controls.Add(Me.cmdButton01)
        Me.Panel1.Location = New System.Drawing.Point(4, 324)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(379, 59)
        Me.Panel1.TabIndex = 23
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.optButton02)
        Me.Panel2.Controls.Add(Me.optButton01)
        Me.Panel2.Controls.Add(Me.optButton00)
        Me.Panel2.Controls.Add(Me.txtField05)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.txtField03)
        Me.Panel2.Controls.Add(Me.txtField06)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Location = New System.Drawing.Point(2, 224)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(379, 98)
        Me.Panel2.TabIndex = 24
        '
        'optButton02
        '
        Me.optButton02.AutoSize = True
        Me.optButton02.Location = New System.Drawing.Point(208, 76)
        Me.optButton02.Name = "optButton02"
        Me.optButton02.Size = New System.Drawing.Size(65, 17)
        Me.optButton02.TabIndex = 9
        Me.optButton02.TabStop = True
        Me.optButton02.Text = "Remove"
        Me.optButton02.UseVisualStyleBackColor = True
        '
        'optButton01
        '
        Me.optButton01.AutoSize = True
        Me.optButton01.Location = New System.Drawing.Point(146, 76)
        Me.optButton01.Name = "optButton01"
        Me.optButton01.Size = New System.Drawing.Size(56, 17)
        Me.optButton01.TabIndex = 8
        Me.optButton01.TabStop = True
        Me.optButton01.Text = "Added"
        Me.optButton01.UseVisualStyleBackColor = True
        '
        'optButton00
        '
        Me.optButton00.AutoSize = True
        Me.optButton00.Location = New System.Drawing.Point(80, 77)
        Me.optButton00.Name = "optButton00"
        Me.optButton00.Size = New System.Drawing.Size(60, 17)
        Me.optButton00.TabIndex = 7
        Me.optButton00.TabStop = True
        Me.optButton00.Text = "Original"
        Me.optButton00.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView1.Location = New System.Drawing.Point(2, 3)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(381, 219)
        Me.DataGridView1.TabIndex = 25
        Me.DataGridView1.TabStop = False
        '
        'frmComboItem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.ggcRetailParams.My.Resources.Resources.mainbackground
        Me.ClientSize = New System.Drawing.Size(385, 385)
        Me.ControlBox = False
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Name = "frmComboItem"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Combo Items"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtField06 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtField05 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtField03 As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents cmdButton03 As System.Windows.Forms.Button
    Friend WithEvents cmdButton02 As System.Windows.Forms.Button
    Friend WithEvents cmdButton01 As System.Windows.Forms.Button
    Friend WithEvents cmdButton00 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents optButton02 As System.Windows.Forms.RadioButton
    Friend WithEvents optButton01 As System.Windows.Forms.RadioButton
    Friend WithEvents optButton00 As System.Windows.Forms.RadioButton
End Class
