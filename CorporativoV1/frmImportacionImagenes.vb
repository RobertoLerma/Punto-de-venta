Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmImportacionImagenes
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             PROCESO DE IMPORTACIÓN DE IMAGENES                                                           *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      MIERCOLES 02 DE JUNIO DE 2004                                                                *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents optAsocCodNue As System.Windows.Forms.RadioButton
    Public WithEvents optAsocCodAnt As System.Windows.Forms.RadioButton
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents lblNumArcNoAbc As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents lblNumArcProcesados As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblNumArcExistentes As System.Windows.Forms.Label
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents chkReemplazarArcExistentes As System.Windows.Forms.CheckBox
    Public WithEvents cmdProcesar As System.Windows.Forms.Button
    Public WithEvents Dir1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents txtRutaDestino As System.Windows.Forms.TextBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents txtRutaOrigen As System.Windows.Forms.TextBox
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox

    Dim mblnSalir As Boolean

    Sub ImportacionImagenes()
        On Error GoTo Err_Renamed
        Dim FileName As String
        Dim cRutaOrigen As String
        Dim cRutaDestino As String
        Dim cNombreArchivo As String
        Dim cExtension As String
        Dim intOrigen As String
        Dim intCodigoAnt As String
        Dim Sql As String
        Dim ArchivoDestino As String
        Dim RsAux As ADODB.Recordset
        Dim nCountImages As Integer
        Dim nCountImagesNotAbc As Integer
        Dim nFilesExistentes As Integer
        Dim Fso As Object
        Dim lArchivo As String
        Dim lRutaArchNoAbc As String
        Dim intCodigoNuevo As String

        Fso = CreateObject("Scripting.FileSystemObject")

        lArchivo = Dir(gstrCorpoDriveLocal & "\Sistema\ImgNoCat", FileAttribute.Directory)
        If lArchivo = "" Then
            MkDir(gstrCorpoDriveLocal & "\Sistema\ImgNoCat")
        End If

        cRutaOrigen = Trim(txtRutaOrigen.Text)
        cRutaDestino = Trim(txtRutaDestino.Text)
        FileName = Dir(cRutaOrigen, FileAttribute.Archive)
        nCountImages = 0
        nFilesExistentes = 0
        nCountImagesNotAbc = 0
        lblNumArcExistentes.Text = CStr(0)
        lblNumArcProcesados.Text = CStr(0)
        lblNumArcNoAbc.Text = CStr(0)
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        lRutaArchNoAbc = gstrCorpoDriveLocal & "\Sistema\ImgNoCat\"
        Do While FileName <> ""
            If FileName <> "." And FileName <> ".." Then
                If (GetAttr(cRutaOrigen & FileName) And FileAttribute.Archive) = FileAttribute.Archive Then 'Validamos que sean archivos y no carpetas
                    'Una vez que tenemos el archivo
                    'Extraemos el nombre y la extensión del archivo
                    cNombreArchivo = Mid(cRutaOrigen & FileName, InStrRev(cRutaOrigen & FileName, "\") + 1, InStrRev(cRutaOrigen & FileName, ".") - (InStrRev(cRutaOrigen & FileName, "\") + 1))
                    cExtension = Mid(cRutaOrigen & FileName, InStrRev(cRutaOrigen & FileName, ".") + 1, Len(cRutaOrigen & FileName) - InStrRev(cRutaOrigen & FileName, "."))
                    'Validamos la Extension del Archivo
                    If UCase(cExtension) = "BMP" Or UCase(cExtension) = "JPG" Or UCase(cExtension) = "GIF" Or UCase(cExtension) = "MPEG" Or UCase(cExtension) = "TIFF" Or UCase(cExtension) = "TIF" Then
                        If optAsocCodAnt.Checked Then
                            intOrigen = (cNombreArchivo)
                            intCodigoAnt = (cNombreArchivo)
                            'Buscamos en el Catalogo de Articulos el Codigo Nuevo del Articulo
                            Sql = "SELECT CodArticulo FROM CatArticulos WHERE OrigenAnt = " & Numerico(intOrigen) & " AND CodigoAnt = " & Numerico(intCodigoAnt)
                        ElseIf optAsocCodNue.Checked Then
                            intCodigoNuevo = Trim(cNombreArchivo)
                            'Buscamos en el Catalogo de Articulos por Codigo Nuevo si existe ese Articulo
                            Sql = "SELECT CodArticulo FROM CatArticulos WHERE CodArticulo = " & CInt(Numerico(intCodigoNuevo))
                        End If
                        ModEstandar.BorraCmd()
                        Cmd.CommandText = "dbo.UP_Select_Datos"
                        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                        RsAux = Cmd.Execute
                        If RsAux.RecordCount > 0 Then
                            ArchivoDestino = RsAux.Fields("CodArticulo").Value & "." & cExtension
                            'Ya que tenemos el Archivo destino
                            'Verificamos si Existe en la ruta destino
                            'UPGRADE_WARNING: Couldn't resolve default property of object Fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If (Fso.FileExists(cRutaDestino & ArchivoDestino)) And chkReemplazarArcExistentes.CheckState = System.Windows.Forms.CheckState.Checked Then
                                'Si existe verificamos el check para saber si se reemplaza
                                'Destruimos el archivo existente
                                Kill(cRutaDestino & ArchivoDestino)
                                'Copiamos el Archivo a la ruta destino
                                FileCopy(cRutaOrigen & FileName, cRutaDestino & ArchivoDestino)
                                nCountImages = nCountImages + 1
                                System.Windows.Forms.Application.DoEvents()
                                lblNumArcProcesados.Text = CStr(nCountImages)
                                'si no existe lo copiamos
                                'UPGRADE_WARNING: Couldn't resolve default property of object Fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            ElseIf Not (Fso.FileExists(cRutaDestino & ArchivoDestino)) Then
                                'Copiamos el Archivo a la ruta destino
                                FileCopy(cRutaOrigen & FileName, cRutaDestino & ArchivoDestino)
                                nCountImages = nCountImages + 1
                                System.Windows.Forms.Application.DoEvents()
                                lblNumArcProcesados.Text = CStr(nCountImages)
                            End If
                        Else
                            ''' si el articulo no existe en el abc entonces se copia con el nombre original ( formato compucaja ) para identificar todos aquellos que no existen en el abc que deberían de existir
                            'Copiamos el Archivo a la ruta destino
                            FileCopy(cRutaOrigen & FileName, lRutaArchNoAbc & cNombreArchivo & "." & cExtension)
                            nCountImagesNotAbc = nCountImagesNotAbc + 1
                            System.Windows.Forms.Application.DoEvents()
                            lblNumArcNoAbc.Text = CStr(nCountImagesNotAbc)
                        End If
                    End If
                    nFilesExistentes = nFilesExistentes + 1
                    lblNumArcExistentes.Text = CStr(nFilesExistentes)
                End If
            End If
            FileName = Dir() 'Tomamos el siguiente archivo
        Loop
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If nCountImages > 0 Then
            MsgBox("Proceso Terminado con Exito", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        End If
        System.Windows.Forms.Application.DoEvents()
Err_Renamed:
        If Err.Number <> 0 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub cmdProcesar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdProcesar.Click
        ImportacionImagenes()
    End Sub

    Private Sub Dir1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Dir1.Change
        If (Dir1.Path) <> "\" Then
            txtRutaOrigen.Text = Dir1.Path & "\"
        Else
            txtRutaOrigen.Text = Dir1.Path
        End If
    End Sub

    Private Sub Dir1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Dir1.Enter
        If (Dir1.Path) <> "\" Then
            txtRutaOrigen.Text = Dir1.Path & "\"
        Else
            txtRutaOrigen.Text = Dir1.Path
        End If
    End Sub

    Private Sub frmImportacionImagenes_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmImportacionImagenes_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmImportacionImagenes_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "txtRutaOrigen" Then
                    ModEstandar.RetrocederTab(Me)
                Else
                    mblnSalir = True
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub frmImportacionImagenes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmImportacionImagenes_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.CentrarForma(Me)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Dir1.Path = gstrCorpoDriveLocal
        txtRutaOrigen.Text = ""
        txtRutaDestino.Text = My.Application.Info.DirectoryPath & "\Sistema\Imagenes\"
        lblNumArcProcesados.Text = CStr(0)
        lblNumArcExistentes.Text = CStr(0)
        chkReemplazarArcExistentes.CheckState = System.Windows.Forms.CheckState.Unchecked
        optAsocCodAnt.Checked = False
        optAsocCodNue.Checked = True
        If (Dir1.Path) <> "\" Then
            txtRutaOrigen.Text = Dir1.Path & "\"
        Else
            txtRutaOrigen.Text = Dir1.Path
        End If
    End Sub

    Private Sub frmImportacionImagenes_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        If mblnSalir Then
            Select Case MsgBox("¿Desea Salir De Este Proceso?", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSalir = False
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmImportacionImagenes_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.optAsocCodNue = New System.Windows.Forms.RadioButton()
        Me.optAsocCodAnt = New System.Windows.Forms.RadioButton()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.lblNumArcNoAbc = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblNumArcProcesados = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblNumArcExistentes = New System.Windows.Forms.Label()
        Me.chkReemplazarArcExistentes = New System.Windows.Forms.CheckBox()
        Me.cmdProcesar = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.Dir1 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.txtRutaDestino = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.txtRutaOrigen = New System.Windows.Forms.TextBox()
        Me.Frame5.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.optAsocCodNue)
        Me.Frame5.Controls.Add(Me.optAsocCodAnt)
        Me.Frame5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame5.Location = New System.Drawing.Point(16, 203)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(305, 50)
        Me.Frame5.TabIndex = 17
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Asociar"
        '
        'optAsocCodNue
        '
        Me.optAsocCodNue.BackColor = System.Drawing.SystemColors.Control
        Me.optAsocCodNue.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAsocCodNue.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.optAsocCodNue.Location = New System.Drawing.Point(165, 20)
        Me.optAsocCodNue.Name = "optAsocCodNue"
        Me.optAsocCodNue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAsocCodNue.Size = New System.Drawing.Size(129, 17)
        Me.optAsocCodNue.TabIndex = 4
        Me.optAsocCodNue.TabStop = True
        Me.optAsocCodNue.Text = "Por código nuevo"
        Me.optAsocCodNue.UseVisualStyleBackColor = False
        '
        'optAsocCodAnt
        '
        Me.optAsocCodAnt.BackColor = System.Drawing.SystemColors.Control
        Me.optAsocCodAnt.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAsocCodAnt.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.optAsocCodAnt.Location = New System.Drawing.Point(14, 20)
        Me.optAsocCodAnt.Name = "optAsocCodAnt"
        Me.optAsocCodAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAsocCodAnt.Size = New System.Drawing.Size(137, 17)
        Me.optAsocCodAnt.TabIndex = 3
        Me.optAsocCodAnt.TabStop = True
        Me.optAsocCodAnt.Text = "Por código anterior"
        Me.optAsocCodAnt.UseVisualStyleBackColor = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.lblNumArcNoAbc)
        Me.Frame4.Controls.Add(Me.Label2)
        Me.Frame4.Controls.Add(Me.Label3)
        Me.Frame4.Controls.Add(Me.lblNumArcProcesados)
        Me.Frame4.Controls.Add(Me.Label1)
        Me.Frame4.Controls.Add(Me.lblNumArcExistentes)
        Me.Frame4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame4.Location = New System.Drawing.Point(16, 120)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(305, 77)
        Me.Frame4.TabIndex = 10
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Total Archivos"
        '
        'lblNumArcNoAbc
        '
        Me.lblNumArcNoAbc.BackColor = System.Drawing.SystemColors.Window
        Me.lblNumArcNoAbc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNumArcNoAbc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNumArcNoAbc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNumArcNoAbc.Location = New System.Drawing.Point(220, 46)
        Me.lblNumArcNoAbc.Name = "lblNumArcNoAbc"
        Me.lblNumArcNoAbc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNumArcNoAbc.Size = New System.Drawing.Size(72, 21)
        Me.lblNumArcNoAbc.TabIndex = 16
        Me.lblNumArcNoAbc.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(220, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(81, 28)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "No existe en catálogo :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(11, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(81, 21)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Procesados :"
        '
        'lblNumArcProcesados
        '
        Me.lblNumArcProcesados.BackColor = System.Drawing.SystemColors.Window
        Me.lblNumArcProcesados.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNumArcProcesados.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNumArcProcesados.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNumArcProcesados.Location = New System.Drawing.Point(99, 45)
        Me.lblNumArcProcesados.Name = "lblNumArcProcesados"
        Me.lblNumArcProcesados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNumArcProcesados.Size = New System.Drawing.Size(65, 21)
        Me.lblNumArcProcesados.TabIndex = 13
        Me.lblNumArcProcesados.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(11, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 21)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Existentes :"
        '
        'lblNumArcExistentes
        '
        Me.lblNumArcExistentes.BackColor = System.Drawing.SystemColors.Window
        Me.lblNumArcExistentes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNumArcExistentes.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNumArcExistentes.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNumArcExistentes.Location = New System.Drawing.Point(99, 21)
        Me.lblNumArcExistentes.Name = "lblNumArcExistentes"
        Me.lblNumArcExistentes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNumArcExistentes.Size = New System.Drawing.Size(65, 21)
        Me.lblNumArcExistentes.TabIndex = 11
        Me.lblNumArcExistentes.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'chkReemplazarArcExistentes
        '
        Me.chkReemplazarArcExistentes.BackColor = System.Drawing.SystemColors.Control
        Me.chkReemplazarArcExistentes.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkReemplazarArcExistentes.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkReemplazarArcExistentes.Location = New System.Drawing.Point(336, 224)
        Me.chkReemplazarArcExistentes.Name = "chkReemplazarArcExistentes"
        Me.chkReemplazarArcExistentes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkReemplazarArcExistentes.Size = New System.Drawing.Size(179, 29)
        Me.chkReemplazarArcExistentes.TabIndex = 5
        Me.chkReemplazarArcExistentes.Text = "Reemplazar Archivos Existentes"
        Me.chkReemplazarArcExistentes.UseVisualStyleBackColor = False
        '
        'cmdProcesar
        '
        Me.cmdProcesar.BackColor = System.Drawing.SystemColors.Control
        Me.cmdProcesar.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdProcesar.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdProcesar.Location = New System.Drawing.Point(517, 224)
        Me.cmdProcesar.Name = "cmdProcesar"
        Me.cmdProcesar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdProcesar.Size = New System.Drawing.Size(100, 25)
        Me.cmdProcesar.TabIndex = 6
        Me.cmdProcesar.Text = "&Procesar"
        Me.cmdProcesar.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.Dir1)
        Me.Frame3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame3.Location = New System.Drawing.Point(336, 16)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(253, 182)
        Me.Frame3.TabIndex = 9
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Seleccionar Ruta de Imagenes Reales"
        '
        'Dir1
        '
        Me.Dir1.BackColor = System.Drawing.SystemColors.Window
        Me.Dir1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Dir1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Dir1.FormattingEnabled = True
        Me.Dir1.IntegralHeight = False
        Me.Dir1.Location = New System.Drawing.Point(6, 17)
        Me.Dir1.Name = "Dir1"
        Me.Dir1.Size = New System.Drawing.Size(240, 156)
        Me.Dir1.TabIndex = 1
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.txtRutaDestino)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(16, 66)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(305, 47)
        Me.Frame2.TabIndex = 8
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Ruta Destino"
        '
        'txtRutaDestino
        '
        Me.txtRutaDestino.AcceptsReturn = True
        Me.txtRutaDestino.BackColor = System.Drawing.SystemColors.Window
        Me.txtRutaDestino.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRutaDestino.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRutaDestino.Location = New System.Drawing.Point(8, 16)
        Me.txtRutaDestino.MaxLength = 0
        Me.txtRutaDestino.Name = "txtRutaDestino"
        Me.txtRutaDestino.ReadOnly = True
        Me.txtRutaDestino.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRutaDestino.Size = New System.Drawing.Size(289, 21)
        Me.txtRutaDestino.TabIndex = 2
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.txtRutaOrigen)
        Me.Frame1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame1.Location = New System.Drawing.Point(16, 16)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(305, 47)
        Me.Frame1.TabIndex = 7
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Ruta de Imagenes Reales"
        '
        'txtRutaOrigen
        '
        Me.txtRutaOrigen.AcceptsReturn = True
        Me.txtRutaOrigen.BackColor = System.Drawing.SystemColors.Window
        Me.txtRutaOrigen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRutaOrigen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRutaOrigen.Location = New System.Drawing.Point(8, 16)
        Me.txtRutaOrigen.MaxLength = 0
        Me.txtRutaOrigen.Name = "txtRutaOrigen"
        Me.txtRutaOrigen.ReadOnly = True
        Me.txtRutaOrigen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRutaOrigen.Size = New System.Drawing.Size(289, 21)
        Me.txtRutaOrigen.TabIndex = 0
        '
        'frmImportacionImagenes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(629, 270)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.chkReemplazarArcExistentes)
        Me.Controls.Add(Me.cmdProcesar)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(206, 101)
        Me.MaximizeBox = False
        Me.Name = "frmImportacionImagenes"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Importar Imagenes"
        Me.Frame5.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

End Class