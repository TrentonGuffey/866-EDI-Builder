<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.btnGo = New System.Windows.Forms.Button()
        Me.txtBeginningLineSet = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtThruDate = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtNGVRackSize = New System.Windows.Forms.TextBox()
        Me.txtFrontRackSize = New System.Windows.Forms.TextBox()
        Me.txtRearRackSize = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtE79RackSize = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnGo
        '
        Me.btnGo.Font = New System.Drawing.Font("Microsoft Sans Serif", 25.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGo.Location = New System.Drawing.Point(555, 196)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(162, 69)
        Me.btnGo.TabIndex = 0
        Me.btnGo.Text = "GO"
        Me.btnGo.UseVisualStyleBackColor = True
        '
        'txtBeginningLineSet
        '
        Me.txtBeginningLineSet.Location = New System.Drawing.Point(158, 19)
        Me.txtBeginningLineSet.Name = "txtBeginningLineSet"
        Me.txtBeginningLineSet.Size = New System.Drawing.Size(96, 20)
        Me.txtBeginningLineSet.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!)
        Me.Label1.Location = New System.Drawing.Point(12, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(125, 18)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Beginning LineSet"
        '
        'txtThruDate
        '
        Me.txtThruDate.Location = New System.Drawing.Point(158, 58)
        Me.txtThruDate.Name = "txtThruDate"
        Me.txtThruDate.Size = New System.Drawing.Size(96, 20)
        Me.txtThruDate.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!)
        Me.Label2.Location = New System.Drawing.Point(12, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 18)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Thru Date"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!)
        Me.Label3.Location = New System.Drawing.Point(12, 107)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 18)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "# of NGV Racks"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!)
        Me.Label4.Location = New System.Drawing.Point(12, 145)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(140, 18)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "# of LT Front Racks"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!)
        Me.Label5.Location = New System.Drawing.Point(12, 181)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(137, 18)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "# of LT Rear Racks"
        '
        'txtNGVRackSize
        '
        Me.txtNGVRackSize.Location = New System.Drawing.Point(222, 108)
        Me.txtNGVRackSize.Name = "txtNGVRackSize"
        Me.txtNGVRackSize.Size = New System.Drawing.Size(32, 20)
        Me.txtNGVRackSize.TabIndex = 8
        '
        'txtFrontRackSize
        '
        Me.txtFrontRackSize.Location = New System.Drawing.Point(222, 145)
        Me.txtFrontRackSize.Name = "txtFrontRackSize"
        Me.txtFrontRackSize.Size = New System.Drawing.Size(32, 20)
        Me.txtFrontRackSize.TabIndex = 9
        '
        'txtRearRackSize
        '
        Me.txtRearRackSize.Location = New System.Drawing.Point(222, 182)
        Me.txtRearRackSize.Name = "txtRearRackSize"
        Me.txtRearRackSize.Size = New System.Drawing.Size(32, 20)
        Me.txtRearRackSize.TabIndex = 10
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!)
        Me.Label6.Location = New System.Drawing.Point(12, 217)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(110, 18)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "# of E79 Racks"
        '
        'txtE79RackSize
        '
        Me.txtE79RackSize.Location = New System.Drawing.Point(222, 218)
        Me.txtE79RackSize.Name = "txtE79RackSize"
        Me.txtE79RackSize.Size = New System.Drawing.Size(32, 20)
        Me.txtE79RackSize.TabIndex = 12
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(772, 300)
        Me.Controls.Add(Me.txtE79RackSize)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtRearRackSize)
        Me.Controls.Add(Me.txtFrontRackSize)
        Me.Controls.Add(Me.txtNGVRackSize)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtThruDate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtBeginningLineSet)
        Me.Controls.Add(Me.btnGo)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "EDI Buider"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnGo As Button
    Friend WithEvents txtBeginningLineSet As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtThruDate As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents txtNGVRackSize As TextBox
    Friend WithEvents txtFrontRackSize As TextBox
    Friend WithEvents txtRearRackSize As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txtE79RackSize As TextBox
End Class
