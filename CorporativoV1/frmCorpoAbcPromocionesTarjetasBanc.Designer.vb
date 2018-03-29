<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCorpoAbcPromocionesTarjetasBanc
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
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents lblActivados As System.Windows.Forms.Label
    Public WithEvents lblSuspendidos As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents cmdActivarSuspender As System.Windows.Forms.Button
    Public WithEvents txtDetalle As System.Windows.Forms.TextBox
    Public WithEvents msgPromocion As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents dbcBanco As System.Windows.Forms.ComboBox
    Public WithEvents txtDescPlan As System.Windows.Forms.Label
    Public WithEvents _lblPromoTarjetas_0 As System.Windows.Forms.Label
    Public WithEvents fraGeneral As System.Windows.Forms.GroupBox
    Public WithEvents lblPromoTarjetas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCorpoAbcPromocionesTarjetasBanc))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDescPlan = New System.Windows.Forms.Label()
        Me.fraGeneral = New System.Windows.Forms.GroupBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.lblActivados = New System.Windows.Forms.Label()
        Me.lblSuspendidos = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdActivarSuspender = New System.Windows.Forms.Button()
        Me.txtDetalle = New System.Windows.Forms.TextBox()
        Me.msgPromocion = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.dbcBanco = New System.Windows.Forms.ComboBox()
        Me._lblPromoTarjetas_0 = New System.Windows.Forms.Label()
        Me.lblPromoTarjetas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.fraGeneral.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.msgPromocion, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPromoTarjetas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtDescPlan
        '
        Me.txtDescPlan.BackColor = System.Drawing.SystemColors.Info
        Me.txtDescPlan.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtDescPlan.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtDescPlan.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.txtDescPlan.Location = New System.Drawing.Point(8, 186)
        Me.txtDescPlan.Name = "txtDescPlan"
        Me.txtDescPlan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDescPlan.Size = New System.Drawing.Size(416, 21)
        Me.txtDescPlan.TabIndex = 5
        Me.txtDescPlan.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.txtDescPlan, "Descripción de Artículos")
        '
        'fraGeneral
        '
        Me.fraGeneral.BackColor = System.Drawing.SystemColors.Control
        Me.fraGeneral.Controls.Add(Me.Frame1)
        Me.fraGeneral.Controls.Add(Me.cmdActivarSuspender)
        Me.fraGeneral.Controls.Add(Me.txtDetalle)
        Me.fraGeneral.Controls.Add(Me.msgPromocion)
        Me.fraGeneral.Controls.Add(Me.dbcBanco)
        Me.fraGeneral.Controls.Add(Me.txtDescPlan)
        Me.fraGeneral.Controls.Add(Me._lblPromoTarjetas_0)
        Me.fraGeneral.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGeneral.Location = New System.Drawing.Point(7, 4)
        Me.fraGeneral.Name = "fraGeneral"
        Me.fraGeneral.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGeneral.Size = New System.Drawing.Size(434, 277)
        Me.fraGeneral.TabIndex = 0
        Me.fraGeneral.TabStop = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.lblActivados)
        Me.Frame1.Controls.Add(Me.lblSuspendidos)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Frame1.Location = New System.Drawing.Point(208, 210)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(121, 57)
        Me.Frame1.TabIndex = 7
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Estatus"
        '
        'lblActivados
        '
        Me.lblActivados.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblActivados.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblActivados.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblActivados.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblActivados.Location = New System.Drawing.Point(8, 16)
        Me.lblActivados.Name = "lblActivados"
        Me.lblActivados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblActivados.Size = New System.Drawing.Size(17, 17)
        Me.lblActivados.TabIndex = 11
        '
        'lblSuspendidos
        '
        Me.lblSuspendidos.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblSuspendidos.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblSuspendidos.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSuspendidos.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblSuspendidos.Location = New System.Drawing.Point(8, 35)
        Me.lblSuspendidos.Name = "lblSuspendidos"
        Me.lblSuspendidos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSuspendidos.Size = New System.Drawing.Size(17, 17)
        Me.lblSuspendidos.TabIndex = 10
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label2.Location = New System.Drawing.Point(35, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(57, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Activos"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label3.Location = New System.Drawing.Point(35, 35)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(81, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Suspendidos"
        '
        'cmdActivarSuspender
        '
        Me.cmdActivarSuspender.BackColor = System.Drawing.SystemColors.Control
        Me.cmdActivarSuspender.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdActivarSuspender.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdActivarSuspender.Location = New System.Drawing.Point(344, 236)
        Me.cmdActivarSuspender.Name = "cmdActivarSuspender"
        Me.cmdActivarSuspender.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdActivarSuspender.Size = New System.Drawing.Size(81, 25)
        Me.cmdActivarSuspender.TabIndex = 6
        Me.cmdActivarSuspender.Text = "Suspender"
        Me.cmdActivarSuspender.UseVisualStyleBackColor = False
        '
        'txtDetalle
        '
        Me.txtDetalle.AcceptsReturn = True
        Me.txtDetalle.BackColor = System.Drawing.SystemColors.Window
        Me.txtDetalle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDetalle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDetalle.Location = New System.Drawing.Point(200, 96)
        Me.txtDetalle.MaxLength = 0
        Me.txtDetalle.Name = "txtDetalle"
        Me.txtDetalle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDetalle.Size = New System.Drawing.Size(65, 20)
        Me.txtDetalle.TabIndex = 3
        Me.txtDetalle.Visible = False
        '
        'msgPromocion
        '
        Me.msgPromocion.DataSource = Nothing
        Me.msgPromocion.Location = New System.Drawing.Point(8, 52)
        Me.msgPromocion.Name = "msgPromocion"
        Me.msgPromocion.OcxState = CType(resources.GetObject("msgPromocion.OcxState"), System.Windows.Forms.AxHost.State)
        Me.msgPromocion.Size = New System.Drawing.Size(416, 130)
        Me.msgPromocion.TabIndex = 4
        '
        'dbcBanco
        '
        Me.dbcBanco.Location = New System.Drawing.Point(56, 20)
        Me.dbcBanco.Name = "dbcBanco"
        Me.dbcBanco.Size = New System.Drawing.Size(237, 21)
        Me.dbcBanco.TabIndex = 2
        '
        '_lblPromoTarjetas_0
        '
        Me._lblPromoTarjetas_0.AutoSize = True
        Me._lblPromoTarjetas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblPromoTarjetas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblPromoTarjetas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPromoTarjetas.SetIndex(Me._lblPromoTarjetas_0, CType(0, Short))
        Me._lblPromoTarjetas_0.Location = New System.Drawing.Point(10, 22)
        Me._lblPromoTarjetas_0.Name = "_lblPromoTarjetas_0"
        Me._lblPromoTarjetas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblPromoTarjetas_0.Size = New System.Drawing.Size(44, 13)
        Me._lblPromoTarjetas_0.TabIndex = 1
        Me._lblPromoTarjetas_0.Text = "Banco :"
        '
        'frmCorpoAbcPromocionesTarjetasBanc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(444, 284)
        Me.Controls.Add(Me.fraGeneral)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(177, 160)
        Me.MaximizeBox = False
        Me.Name = "frmCorpoAbcPromocionesTarjetasBanc"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "ABC  a  Promociones Tarjetas Bancarias"
        Me.fraGeneral.ResumeLayout(False)
        Me.fraGeneral.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        CType(Me.msgPromocion, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPromoTarjetas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
End Class
