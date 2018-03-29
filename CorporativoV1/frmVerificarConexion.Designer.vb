Imports Microsoft.VisualBasic.Compatibility

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmVerificarConexion
#Region "Windows Form Designer generated code "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        'InitializeComponent()
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
    'Required by the Windows Form Designer


    Public WithEvents TxtNomServidor As TextBox
    Public WithEvents TxtBDPrincipal As TextBox
    Public WithEvents _LblEtiqueta_1 As Label
    Public WithEvents LblEtiqueta As VB6.LabelArray
    Private components As System.ComponentModel.IContainer
    Public WithEvents ToolTip1 As ToolTip
    Public WithEvents _LblEtiqueta_0 As Label     'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.TxtNomServidor = New System.Windows.Forms.TextBox()
        Me.TxtBDPrincipal = New System.Windows.Forms.TextBox()
        Me._LblEtiqueta_1 = New System.Windows.Forms.Label()
        Me._LblEtiqueta_0 = New System.Windows.Forms.Label()
        Me.LblEtiqueta = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.LblEtiqueta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TxtNomServidor
        '
        Me.TxtNomServidor.AcceptsReturn = True
        Me.TxtNomServidor.BackColor = System.Drawing.SystemColors.Window
        Me.TxtNomServidor.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtNomServidor.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtNomServidor.Location = New System.Drawing.Point(328, 111)
        Me.TxtNomServidor.MaxLength = 80
        Me.TxtNomServidor.Name = "TxtNomServidor"
        Me.TxtNomServidor.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtNomServidor.Size = New System.Drawing.Size(127, 22)
        Me.TxtNomServidor.TabIndex = 0
        '
        'TxtBDPrincipal
        '
        Me.TxtBDPrincipal.AcceptsReturn = True
        Me.TxtBDPrincipal.BackColor = System.Drawing.SystemColors.Window
        Me.TxtBDPrincipal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtBDPrincipal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtBDPrincipal.Location = New System.Drawing.Point(328, 189)
        Me.TxtBDPrincipal.MaxLength = 80
        Me.TxtBDPrincipal.Name = "TxtBDPrincipal"
        Me.TxtBDPrincipal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtBDPrincipal.Size = New System.Drawing.Size(127, 22)
        Me.TxtBDPrincipal.TabIndex = 1
        '
        '_LblEtiqueta_1
        '
        Me._LblEtiqueta_1.AutoSize = True
        Me._LblEtiqueta_1.BackColor = System.Drawing.SystemColors.Control
        Me._LblEtiqueta_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._LblEtiqueta_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._LblEtiqueta_1.Location = New System.Drawing.Point(355, 81)
        Me._LblEtiqueta_1.Name = "_LblEtiqueta_1"
        Me._LblEtiqueta_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._LblEtiqueta_1.Size = New System.Drawing.Size(65, 17)
        Me._LblEtiqueta_1.TabIndex = 4
        Me._LblEtiqueta_1.Text = "Servidor "
        '
        '_LblEtiqueta_0
        '
        Me._LblEtiqueta_0.AutoSize = True
        Me._LblEtiqueta_0.BackColor = System.Drawing.SystemColors.Control
        Me._LblEtiqueta_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._LblEtiqueta_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._LblEtiqueta_0.Location = New System.Drawing.Point(335, 158)
        Me._LblEtiqueta_0.Name = "_LblEtiqueta_0"
        Me._LblEtiqueta_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._LblEtiqueta_0.Size = New System.Drawing.Size(105, 17)
        Me._LblEtiqueta_0.TabIndex = 3
        Me._LblEtiqueta_0.Text = "Base de Datos "
        '
        'frmVerificarConexion
        '
        Me.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.conexion
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(823, 418)
        Me.Controls.Add(Me._LblEtiqueta_0)
        Me.Controls.Add(Me._LblEtiqueta_1)
        Me.Controls.Add(Me.TxtBDPrincipal)
        Me.Controls.Add(Me.TxtNomServidor)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmVerificarConexion"
        Me.Text = "frmVerificarConexion"
        CType(Me.LblEtiqueta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class