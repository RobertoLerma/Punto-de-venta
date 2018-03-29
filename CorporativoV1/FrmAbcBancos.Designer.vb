<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmAbcBancos
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes. 
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'This form is an MDI child.
        'This code simulates the VB6 
        ' functionality of automatically
        ' loading and showing an MDI
        ' child's parent.
        'Me.MdiParent = MDIMenuPrincipalCorpo
        'MDIMenuPrincipalCorpo.Show()
        'Form_Initialize_Renamed()
    End Sub
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms 
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkSucursal As System.Windows.Forms.CheckBox
    Public WithEvents chkBancoInterno As System.Windows.Forms.CheckBox
    Public WithEvents txtCodBanco As System.Windows.Forms.TextBox
    Public WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Public WithEvents _lblBancos_1 As System.Windows.Forms.Label
    Public WithEvents _lblBancos_0 As System.Windows.Forms.Label
    Public WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Public WithEvents lblBancos As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents btnLimpiar As System.Windows.Forms.Button
    Friend WithEvents btnEliminar As System.Windows.Forms.Button
    Friend WithEvents btnGuardar As System.Windows.Forms.Button
    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.fraGeneral = New System.Windows.Forms.GroupBox()
        Me.chkSucursal = New System.Windows.Forms.CheckBox()
        Me.chkBancoInterno = New System.Windows.Forms.CheckBox()
        Me.txtCodBanco = New System.Windows.Forms.TextBox()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me._lblBancos_1 = New System.Windows.Forms.Label()
        Me._lblBancos_0 = New System.Windows.Forms.Label()
        Me.lblBancos = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnLimpiar = New System.Windows.Forms.Button()
        Me.btnEliminar = New System.Windows.Forms.Button()
        Me.btnGuardar = New System.Windows.Forms.Button()
        Me.fraGeneral.SuspendLayout()
        CType(Me.lblBancos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraGeneral
        '
        Me.fraGeneral.BackColor = System.Drawing.SystemColors.Control
        Me.fraGeneral.Controls.Add(Me.chkSucursal)
        Me.fraGeneral.Controls.Add(Me.chkBancoInterno)
        Me.fraGeneral.Controls.Add(Me.txtCodBanco)
        Me.fraGeneral.Controls.Add(Me.txtDescripcion)
        Me.fraGeneral.Controls.Add(Me._lblBancos_1)
        Me.fraGeneral.Controls.Add(Me._lblBancos_0)
        Me.fraGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGeneral.Location = New System.Drawing.Point(8, 7)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGeneral.Size = New System.Drawing.Size(432, 181)
        Me.fraGeneral.TabIndex = 4
        Me.fraGeneral.TabStop = False
        Me.ToolTip1.SetToolTip(Me.fraGeneral, "Descripción")
        '
        'chkSucursal
        '
        Me.chkSucursal.BackColor = System.Drawing.SystemColors.Control
        Me.chkSucursal.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSucursal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSucursal.Location = New System.Drawing.Point(16, 130)
        Me.chkSucursal.Name = "chkSucursal"
        Me.chkSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSucursal.Size = New System.Drawing.Size(117, 21)
        Me.chkSucursal.TabIndex = 3
        Me.chkSucursal.Text = "Sucursal"
        Me.chkSucursal.UseVisualStyleBackColor = False
        '
        'chkBancoInterno
        '
        Me.chkBancoInterno.BackColor = System.Drawing.SystemColors.Control
        Me.chkBancoInterno.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBancoInterno.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBancoInterno.Location = New System.Drawing.Point(16, 93)
        Me.chkBancoInterno.Name = "chkBancoInterno"
        Me.chkBancoInterno.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBancoInterno.Size = New System.Drawing.Size(135, 21)
        Me.chkBancoInterno.TabIndex = 2
        Me.chkBancoInterno.Text = "Banco Interno"
        Me.chkBancoInterno.UseVisualStyleBackColor = False
        '
        'txtCodBanco
        '
        Me.txtCodBanco.AcceptsReturn = True
        Me.txtCodBanco.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodBanco.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodBanco.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodBanco.Location = New System.Drawing.Point(80, 28)
        Me.txtCodBanco.MaxLength = 3
        Me.txtCodBanco.Name = "txtCodBanco"
        Me.txtCodBanco.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodBanco.Size = New System.Drawing.Size(71, 22)
        Me.txtCodBanco.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtCodBanco, "Código del Banco")
        '
        'txtDescripcion
        '
        Me.txtDescripcion.AcceptsReturn = True
        Me.txtDescripcion.BackColor = System.Drawing.SystemColors.Window
        Me.txtDescripcion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDescripcion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDescripcion.Location = New System.Drawing.Point(110, 59)
        Me.txtDescripcion.MaxLength = 40
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescripcion.Size = New System.Drawing.Size(265, 22)
        Me.txtDescripcion.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtDescripcion, "Descripción del Banco")
        '
        '_lblBancos_1
        '
        Me._lblBancos_1.AutoSize = True
        Me._lblBancos_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblBancos_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBancos_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBancos.SetIndex(Me._lblBancos_1, CType(1, Short))
        Me._lblBancos_1.Location = New System.Drawing.Point(18, 59)
        Me._lblBancos_1.Name = "_lblBancos_1"
        Me._lblBancos_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBancos_1.Size = New System.Drawing.Size(86, 17)
        Me._lblBancos_1.TabIndex = 6
        Me._lblBancos_1.Text = "Descripción:"
        '
        '_lblBancos_0
        '
        Me._lblBancos_0.AutoSize = True
        Me._lblBancos_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblBancos_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBancos_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBancos.SetIndex(Me._lblBancos_0, CType(0, Short))
        Me._lblBancos_0.Location = New System.Drawing.Point(18, 28)
        Me._lblBancos_0.Name = "_lblBancos_0"
        Me._lblBancos_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBancos_0.Size = New System.Drawing.Size(56, 17)
        Me._lblBancos_0.TabIndex = 5
        Me._lblBancos_0.Text = "Código:"
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Location = New System.Drawing.Point(304, 207)
        Me.btnLimpiar.Margin = New System.Windows.Forms.Padding(4)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(124, 43)
        Me.btnLimpiar.TabIndex = 62
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnEliminar
        '
        Me.btnEliminar.Location = New System.Drawing.Point(161, 207)
        Me.btnEliminar.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEliminar.Name = "btnEliminar"
        Me.btnEliminar.Size = New System.Drawing.Size(124, 43)
        Me.btnEliminar.TabIndex = 61
        Me.btnEliminar.Text = "Eliminar"
        Me.btnEliminar.UseVisualStyleBackColor = True
        '
        'btnGuardar
        '
        Me.btnGuardar.Location = New System.Drawing.Point(17, 207)
        Me.btnGuardar.Margin = New System.Windows.Forms.Padding(4)
        Me.btnGuardar.Name = "btnGuardar"
        Me.btnGuardar.Size = New System.Drawing.Size(124, 43)
        Me.btnGuardar.TabIndex = 60
        Me.btnGuardar.Text = "Guardar"
        Me.btnGuardar.UseVisualStyleBackColor = True
        '
        'FrmAbcBancos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(474, 270)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnEliminar)
        Me.Controls.Add(Me.btnGuardar)
        Me.Controls.Add(Me.fraGeneral)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(177, 160)
        Me.MaximizeBox = False
        Me.Name = "FrmAbcBancos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC a Bancos"
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        CType(Me.lblBancos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


End Class
