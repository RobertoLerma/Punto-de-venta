<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAcercaDe
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAcercaDe))
        Me.lblUltimaCompilacion = New System.Windows.Forms.Label()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.lblDisclaimer = New System.Windows.Forms.Label()
        Me.Image1 = New System.Windows.Forms.PictureBox()
        Me.lblfechaCompilacion = New System.Windows.Forms.Label()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblUltimaCompilacion
        '
        Me.lblUltimaCompilacion.BackColor = System.Drawing.SystemColors.Control
        Me.lblUltimaCompilacion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblUltimaCompilacion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblUltimaCompilacion.Location = New System.Drawing.Point(157, 115)
        Me.lblUltimaCompilacion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblUltimaCompilacion.Name = "lblUltimaCompilacion"
        Me.lblUltimaCompilacion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblUltimaCompilacion.Size = New System.Drawing.Size(110, 14)
        Me.lblUltimaCompilacion.TabIndex = 10
        Me.lblUltimaCompilacion.Text = "Última compilación:"
        '
        'lblDescription
        '
        Me.lblDescription.BackColor = System.Drawing.SystemColors.Control
        Me.lblDescription.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescription.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblDescription.Location = New System.Drawing.Point(155, 58)
        Me.lblDescription.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescription.Size = New System.Drawing.Size(227, 38)
        Me.lblDescription.TabIndex = 5
        Me.lblDescription.Text = "Sistema de Control de Joyería para Joyería y Regalos S.A. de C.V."
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
        Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTitle.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblTitle.Location = New System.Drawing.Point(155, 11)
        Me.lblTitle.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTitle.Size = New System.Drawing.Size(131, 17)
        Me.lblTitle.TabIndex = 8
        Me.lblTitle.Text = "Título de la aplicación"
        '
        'lblVersion
        '
        Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblVersion.Location = New System.Drawing.Point(155, 40)
        Me.lblVersion.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(70, 12)
        Me.lblVersion.TabIndex = 9
        Me.lblVersion.Text = "Versión"
        '
        'lblDisclaimer
        '
        Me.lblDisclaimer.BackColor = System.Drawing.SystemColors.Control
        Me.lblDisclaimer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDisclaimer.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblDisclaimer.Location = New System.Drawing.Point(203, 145)
        Me.lblDisclaimer.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblDisclaimer.Name = "lblDisclaimer"
        Me.lblDisclaimer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDisclaimer.Size = New System.Drawing.Size(152, 17)
        Me.lblDisclaimer.TabIndex = 6
        Me.lblDisclaimer.Text = "Monterrey, Nuevo Leon, 2018"
        Me.lblDisclaimer.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Image1
        '
        Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
        Me.Image1.Location = New System.Drawing.Point(16, 25)
        Me.Image1.Margin = New System.Windows.Forms.Padding(2)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(134, 111)
        Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Image1.TabIndex = 7
        Me.Image1.TabStop = False
        '
        'lblfechaCompilacion
        '
        Me.lblfechaCompilacion.AutoSize = True
        Me.lblfechaCompilacion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblfechaCompilacion.Location = New System.Drawing.Point(257, 115)
        Me.lblfechaCompilacion.Name = "lblfechaCompilacion"
        Me.lblfechaCompilacion.Size = New System.Drawing.Size(0, 13)
        Me.lblfechaCompilacion.TabIndex = 11
        '
        'frmAcercaDe
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(392, 180)
        Me.Controls.Add(Me.lblfechaCompilacion)
        Me.Controls.Add(Me.Image1)
        Me.Controls.Add(Me.lblUltimaCompilacion)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.lblDisclaimer)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmAcercaDe"
        Me.Text = "frmAcercaDe"
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Public WithEvents Image1 As PictureBox
    Public WithEvents lblUltimaCompilacion As Label
    Public WithEvents lblDescription As Label
    Public WithEvents lblTitle As Label
    Public WithEvents lblVersion As Label
    Public WithEvents lblDisclaimer As Label
    Friend WithEvents lblfechaCompilacion As Label
End Class
