Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmBarraDesplazamiento
    Inherits System.Windows.Forms.Form

    Public components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents PrgBarra As System.Windows.Forms.ProgressBar
    Public WithEvents Frame As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As System.Windows.Forms.Label

    Private Sub frmBarraDesplazamiento_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        Top = VB6.TwipsToPixelsY(3775)
        Left = VB6.TwipsToPixelsX(5600)
        Icono(Me, MDIMenuPrincipalCorpo)
        'Caption = mstrProcesoInv
    End Sub

    Private Sub frmBarraDesplazamiento_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Dim varF As Object
        'Me = Nothing
        IsNothing(Me)
        'varF = ObtenerForma(Me.Tag)
        'If Not varF Is Nothing Then varF.ZOrder()
        '   Select Case UCase(Trim(Me.Tag))
        '          Case "FRMINVCAPTURAINVFISICO"
        ''               frmInvCapturaInvFisico.ZOrder
        '          Case "FRMINVANALISISCOMPARATIVO"
        '               'frmInvAnalisisComparativo.ZOrder
        '   End Select
    End Sub


    Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBarraDesplazamiento))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.Frame = New System.Windows.Forms.GroupBox
        Me.PrgBarra = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.Frame.SuspendLayout()
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.ClientSize = New System.Drawing.Size(294, 88)
        Me.Location = New System.Drawing.Point(416, 349)
        Me.MaximizeBox = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.MinimizeBox = False
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ControlBox = True
        Me.Enabled = True
        Me.KeyPreview = False
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = True
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmBarraDesplazamiento"
        Me.Frame.Size = New System.Drawing.Size(276, 52)
        Me.Frame.Location = New System.Drawing.Point(10, 26)
        Me.Frame.TabIndex = 0
        Me.Frame.BackColor = System.Drawing.SystemColors.Control
        Me.Frame.Enabled = True
        Me.Frame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame.Visible = True
        Me.Frame.Name = "Frame"
        Me.PrgBarra.Size = New System.Drawing.Size(258, 21)
        Me.PrgBarra.Location = New System.Drawing.Point(9, 19)
        Me.PrgBarra.TabIndex = 1
        Me.PrgBarra.Name = "PrgBarra"
        Me.Label1.Text = "Espere un momento...    cargando  información"
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(64, 64, 64)
        Me.Label1.Size = New System.Drawing.Size(268, 17)
        Me.Label1.Location = New System.Drawing.Point(17, 7)
        Me.Label1.TabIndex = 2
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Enabled = True
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.UseMnemonic = True
        Me.Label1.Visible = True
        Me.Label1.AutoSize = False
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Label1.Name = "Label1"
        Me.Controls.Add(Frame)
        Me.Controls.Add(Label1)
        Me.Frame.Controls.Add(PrgBarra)
        Me.Frame.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

End Class