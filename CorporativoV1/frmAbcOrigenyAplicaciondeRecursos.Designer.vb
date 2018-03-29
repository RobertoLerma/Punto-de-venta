<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmAbcOrigenyAplicaciondeRecursos
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me._optTipoAplicacion_1 = New System.Windows.Forms.RadioButton()
        Me._optTipoAplicacion_0 = New System.Windows.Forms.RadioButton()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me.txtCodigo = New System.Windows.Forms.TextBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnSalir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.Silver
        Me.Frame1.Controls.Add(Me.Frame2)
        Me.Frame1.Controls.Add(Me.txtDescripcion)
        Me.Frame1.Controls.Add(Me.txtCodigo)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(13, 14)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(405, 110)
        Me.Frame1.TabIndex = 5
        Me.Frame1.TabStop = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.Color.Silver
        Me.Frame2.Controls.Add(Me._optTipoAplicacion_1)
        Me.Frame2.Controls.Add(Me._optTipoAplicacion_0)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(166, 10)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(224, 48)
        Me.Frame2.TabIndex = 7
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Aplicación"
        '
        '_optTipoAplicacion_1
        '
        Me._optTipoAplicacion_1.BackColor = System.Drawing.Color.Silver
        Me._optTipoAplicacion_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoAplicacion_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optTipoAplicacion_1.Location = New System.Drawing.Point(151, 16)
        Me._optTipoAplicacion_1.Name = "_optTipoAplicacion_1"
        Me._optTipoAplicacion_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoAplicacion_1.Size = New System.Drawing.Size(67, 26)
        Me._optTipoAplicacion_1.TabIndex = 2
        Me._optTipoAplicacion_1.TabStop = True
        Me._optTipoAplicacion_1.Text = "Salida"
        Me._optTipoAplicacion_1.UseVisualStyleBackColor = False
        '
        '_optTipoAplicacion_0
        '
        Me._optTipoAplicacion_0.BackColor = System.Drawing.Color.Silver
        Me._optTipoAplicacion_0.Checked = True
        Me._optTipoAplicacion_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optTipoAplicacion_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optTipoAplicacion_0.Location = New System.Drawing.Point(75, 16)
        Me._optTipoAplicacion_0.Name = "_optTipoAplicacion_0"
        Me._optTipoAplicacion_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optTipoAplicacion_0.Size = New System.Drawing.Size(66, 26)
        Me._optTipoAplicacion_0.TabIndex = 1
        Me._optTipoAplicacion_0.TabStop = True
        Me._optTipoAplicacion_0.Text = "Entrada"
        Me._optTipoAplicacion_0.UseVisualStyleBackColor = False
        '
        'txtDescripcion
        '
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescripcion.Location = New System.Drawing.Point(89, 74)
        Me.txtDescripcion.MaxLength = 40
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(300, 20)
        Me.txtDescripcion.TabIndex = 3
        '
        'txtCodigo
        '
        Me.txtCodigo.AcceptsReturn = True
        Me.txtCodigo.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodigo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodigo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodigo.Location = New System.Drawing.Point(89, 48)
        Me.txtCodigo.MaxLength = 4
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodigo.Size = New System.Drawing.Size(49, 20)
        Me.txtCodigo.TabIndex = 0
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.Color.Silver
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(10, 81)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(76, 12)
        Me._Label1_1.TabIndex = 6
        Me._Label1_1.Text = "Descripción :"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.Color.Silver
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(10, 56)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(50, 15)
        Me._Label1_0.TabIndex = 5
        Me._Label1_0.Text = "Código :"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gainsboro
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Frame1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(433, 217)
        Me.Panel1.TabIndex = 6
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Silver
        Me.Panel3.Controls.Add(Me.btnSalir)
        Me.Panel3.Controls.Add(Me.btnBuscar)
        Me.Panel3.Controls.Add(Me.btnGuardar)
        Me.Panel3.Controls.Add(Me.btnLimpiar)
        Me.Panel3.Controls.Add(Me.btnEliminar)
        Me.Panel3.Location = New System.Drawing.Point(13, 129)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(405, 74)
        Me.Panel3.TabIndex = 72
        '
        'btnSalir
        '
        Me.btnSalir.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.salir
        Me.btnSalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalir.Location = New System.Drawing.Point(208, 14)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(50, 42)
        Me.btnSalir.TabIndex = 70
        Me.btnSalir.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.buscar
        Me.btnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBuscar.Location = New System.Drawing.Point(160, 14)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(50, 42)
        Me.btnBuscar.TabIndex = 67
        Me.btnBuscar.Text = " "
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.grabar
        Me.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnGuardar.Location = New System.Drawing.Point(11, 14)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(50, 42)
        Me.btnGuardar.TabIndex = 64
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.nuevo
        Me.btnLimpiar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLimpiar.Location = New System.Drawing.Point(110, 14)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(50, 42)
        Me.btnLimpiar.TabIndex = 66
        Me.btnLimpiar.Text = " "
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.Eliminar
        Me.btnEliminar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnEliminar.Location = New System.Drawing.Point(61, 14)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(50, 42)
        Me.btnEliminar.TabIndex = 65
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'frmAbcOrigenyAplicaciondeRecursos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.fondos2
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(459, 241)
        Me.Controls.Add(Me.Panel1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmAbcOrigenyAplicaciondeRecursos"
        Me.Text = "frmAbcOrigenyAplicaciondeRecursos"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Public WithEvents Frame1 As GroupBox
    Public WithEvents Frame2 As GroupBox
    Public WithEvents _optTipoAplicacion_1 As RadioButton
    Public WithEvents _optTipoAplicacion_0 As RadioButton
    Public WithEvents txtDescripcion As TextBox
    Public WithEvents txtCodigo As TextBox
    Public WithEvents _Label1_1 As Label
    Public WithEvents _Label1_0 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnSalir As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnGuardar As Button
    Friend WithEvents btnLimpiar As Button
    Friend WithEvents btnEliminar As Button
End Class
