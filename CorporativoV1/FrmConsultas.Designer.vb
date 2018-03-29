<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmConsultas
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

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents Flexdet As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid


    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        'Me.SuspendLayout()
        ''
        ''FrmConsultas
        ''
        'Me.ClientSize = New System.Drawing.Size(520, 240)
        'Me.Name = "FrmConsultas"
        'Me.ResumeLayout(False)

        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmConsultas))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip
        Me.Flexdet = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        CType(Me.Flexdet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Text = " "
        Me.ClientSize = New System.Drawing.Size(387, 188)
        Me.Location = New System.Drawing.Point(196, 148)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ControlBox = True
        Me.Enabled = True
        Me.KeyPreview = False
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "FrmConsultas"
        Flexdet.OcxState = CType(resources.GetObject("Flexdet.OcxState"), System.Windows.Forms.AxHost.State)
        Me.Flexdet.Size = New System.Drawing.Size(369, 177)
        Me.Flexdet.Location = New System.Drawing.Point(8, 5)
        Me.Flexdet.TabIndex = 0
        Me.Flexdet.Name = "Flexdet"
        Me.Controls.Add(Flexdet)
        CType(Me.Flexdet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()


    End Sub
End Class
