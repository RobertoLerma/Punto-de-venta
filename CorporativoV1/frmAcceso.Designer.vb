<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class FrmAcceso
#Region "Windows Form Designer generated code "
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
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents btnExit As Button
    Public WithEvents PictureBox1 As PictureBox
    Public WithEvents TxtUsuario As TextBox
    Public WithEvents TxtClave As TextBox
    Public WithEvents btnLogin As PictureBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.TxtUsuario = New System.Windows.Forms.TextBox()
        Me.TxtClave = New System.Windows.Forms.TextBox()
        Me.btnLogin = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.btnLogin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(154, Byte), Integer), CType(CType(133, Byte), Integer))
        Me.btnExit.Font = New System.Drawing.Font("Arial", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.btnExit.Location = New System.Drawing.Point(664, 422)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(69, 32)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "Salir"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'TxtUsuario
        '
        Me.TxtUsuario.BackColor = System.Drawing.Color.White
        Me.TxtUsuario.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtUsuario.Font = New System.Drawing.Font("Microsoft JhengHei UI Light", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtUsuario.ForeColor = System.Drawing.Color.Black
        Me.TxtUsuario.Location = New System.Drawing.Point(117, 149)
        Me.TxtUsuario.Multiline = True
        Me.TxtUsuario.Name = "TxtUsuario"
        Me.TxtUsuario.Size = New System.Drawing.Size(115, 26)
        Me.TxtUsuario.TabIndex = 6
        Me.TxtUsuario.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TxtClave
        '
        Me.TxtClave.BackColor = System.Drawing.Color.White
        Me.TxtClave.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtClave.Font = New System.Drawing.Font("Microsoft JhengHei UI Light", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtClave.ForeColor = System.Drawing.Color.Black
        Me.TxtClave.Location = New System.Drawing.Point(114, 232)
        Me.TxtClave.Name = "TxtClave"
        Me.TxtClave.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtClave.Size = New System.Drawing.Size(123, 21)
        Me.TxtClave.TabIndex = 7
        Me.TxtClave.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnLogin
        '
        Me.btnLogin.BackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(154, Byte), Integer), CType(CType(133, Byte), Integer))
        Me.btnLogin.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.Entrar1
        Me.btnLogin.Location = New System.Drawing.Point(76, 340)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(160, 42)
        Me.btnLogin.TabIndex = 8
        Me.btnLogin.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.BackgroundImage = Global.CorporativoV1.My.Resources.Resources.inicio
        Me.PictureBox1.Location = New System.Drawing.Point(0, -1)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(737, 441)
        Me.PictureBox1.TabIndex = 5
        Me.PictureBox1.TabStop = False
        '
        'FrmAcceso
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit
        Me.BackColor = System.Drawing.Color.Gray
        Me.ClientSize = New System.Drawing.Size(770, 508)
        Me.ControlBox = False
        Me.Controls.Add(Me.TxtUsuario)
        Me.Controls.Add(Me.TxtClave)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnLogin)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "FrmAcceso"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.TopMost = True
        Me.TransparencyKey = System.Drawing.Color.Gray
        CType(Me.btnLogin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class