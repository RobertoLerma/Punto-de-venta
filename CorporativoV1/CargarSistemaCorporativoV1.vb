'**********************************************************************************************************************'
'*PROGRAMA: CARGAR SISTEMA CORPORATIVO JOYERIA RAMOS
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports System.IO
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class CargarSistemaCorporativoV1
    Inherits System.Windows.Forms.Form
    Dim dtConfiguracion As New DataTable
    Private WithEvents progressBar1 As ProgressBar
    Friend WithEvents lblBarra1 As Label
    Dim dsConfiguracion As New DataSet
    Dim frmAcceso1 As FrmAcceso = New FrmAcceso()
    Dim frmVerificarConexion1 As frmVerificarConexion = New frmVerificarConexion()


    Private Sub CargarSistema_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        ModEstandar.CentrarForma(Me)
        'Dim processes() As Process
        'Dim contador As Integer

        'contador = Process.GetProcessesByName("Midas").Length
        'If contador > 1 Then
        'End
        'End If

        'Dim Login As New procLogin
        'Me.Hide()
        'Login.Show()


        Try
            'si existe el archivo 
            'entra a forma de acceso y esconde la de verificarconexion
            'sino
            'abre forma de verficarconexion para crear el archivo y 
            'esconde forma de acceso 

            'If (ArchivoTxt.FileExists("C:\\Users\\Consultor_Vitek\\Desktop\\PROYECTO CORPORATIVO\\CORPORATIVO Y JOYERIA\\CODIGO ANGEL WHA\\CorporativoV1\\CorporativoV1\\Sistema\\CJoyeria.Txt")) Then
            If System.IO.File.Exists(rutaArchivoTxt) Then
                Me.Hide()
                frmVerificarConexion1.Hide()
                frmAcceso1.Show()
                progressBar1.Visible = True
                lblBarra1.Visible = True
                ModEstandar.CentrarForma(frmAcceso1)
            Else
                Me.Hide()
                'frmAcceso1.Close()
                frmVerificarConexion1.Show()
                progressBar1.Visible = True
                lblBarra1.Visible = True
                ModEstandar.CentrarForma(frmVerificarConexion1)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub



#Region " Código generado por el Diseñador de Windows Forms "
    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub
    Private components As System.ComponentModel.IContainer
    Friend WithEvents PicMidas As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CargarSistemaCorporativoV1))
        Me.PicMidas = New System.Windows.Forms.PictureBox()
        Me.progressBar1 = New System.Windows.Forms.ProgressBar()
        Me.lblBarra1 = New System.Windows.Forms.Label()
        CType(Me.PicMidas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PicMidas
        '
        Me.PicMidas.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PicMidas.Image = Global.CorporativoV1.My.Resources.Resources.vitek11
        Me.PicMidas.Location = New System.Drawing.Point(0, 0)
        Me.PicMidas.Name = "PicMidas"
        Me.PicMidas.Size = New System.Drawing.Size(643, 366)
        Me.PicMidas.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicMidas.TabIndex = 0
        Me.PicMidas.TabStop = False
        '
        'progressBar1
        '
        Me.progressBar1.Location = New System.Drawing.Point(123, 324)
        Me.progressBar1.Name = "progressBar1"
        Me.progressBar1.Size = New System.Drawing.Size(407, 30)
        Me.progressBar1.TabIndex = 19
        '
        'lblBarra1
        '
        Me.lblBarra1.AutoSize = True
        Me.lblBarra1.Font = New System.Drawing.Font("Arial Narrow", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBarra1.Location = New System.Drawing.Point(270, 296)
        Me.lblBarra1.Name = "lblBarra1"
        Me.lblBarra1.Size = New System.Drawing.Size(105, 25)
        Me.lblBarra1.TabIndex = 20
        Me.lblBarra1.Text = "Cargando..."
        Me.lblBarra1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'CargarSistemaCorporativoV1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(643, 366)
        Me.Controls.Add(Me.lblBarra1)
        Me.Controls.Add(Me.progressBar1)
        Me.Controls.Add(Me.PicMidas)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "CargarSistemaCorporativoV1"
        Me.Text = "CargarSistema"
        CType(Me.PicMidas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region

End Class