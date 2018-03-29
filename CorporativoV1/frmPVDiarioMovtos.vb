Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmPVDiarioMovtos
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    'Programa: Reporte de Movimientos Diarios
    'Autor: Rosaura Torres López
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chk_CodigoAnt As System.Windows.Forms.CheckBox
    Public WithEvents dtpFechaReporte As System.Windows.Forms.DateTimePicker
    Public WithEvents Frame5 As System.Windows.Forms.Panel
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents dbcCaja As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim mblnNuevo As Integer
    Dim mblnTecleoFechaI As Boolean
    Dim msglTiempoCambioI As Single '''Variable para controlar el cambio en el date picker de fecha
    Dim mintCodSucursal As Integer
    Dim mintCodCaja As Integer


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.chk_CodigoAnt = New System.Windows.Forms.CheckBox()
        Me.Frame5 = New System.Windows.Forms.Panel()
        Me.dtpFechaReporte = New System.Windows.Forms.DateTimePicker()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me.dbcCaja = New System.Windows.Forms.ComboBox()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Frame1.SuspendLayout()
        Me.Frame5.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Integer))
        Me._Label1_1.Location = New System.Drawing.Point(7, 50)
        Me._Label1_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(55, 14)
        Me._Label1_1.TabIndex = 7
        Me._Label1_1.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_1, "Nombre de la Farmacia Actual")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.chk_CodigoAnt)
        Me.Frame1.Controls.Add(Me.Frame5)
        Me.Frame1.Controls.Add(Me.dbcSucursal)
        Me.Frame1.Controls.Add(Me.dbcCaja)
        Me.Frame1.Controls.Add(Me._lblVentas_5)
        Me.Frame1.Controls.Add(Me._lblVentas_1)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(7, 2)
        Me.Frame1.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(2)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(242, 128)
        Me.Frame1.TabIndex = 4
        Me.Frame1.TabStop = False
        '
        'chk_CodigoAnt
        '
        Me.chk_CodigoAnt.BackColor = System.Drawing.SystemColors.Control
        Me.chk_CodigoAnt.Cursor = System.Windows.Forms.Cursors.Default
        Me.chk_CodigoAnt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chk_CodigoAnt.Location = New System.Drawing.Point(14, 104)
        Me.chk_CodigoAnt.Margin = New System.Windows.Forms.Padding(2)
        Me.chk_CodigoAnt.Name = "chk_CodigoAnt"
        Me.chk_CodigoAnt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chk_CodigoAnt.Size = New System.Drawing.Size(141, 17)
        Me.chk_CodigoAnt.TabIndex = 3
        Me.chk_CodigoAnt.Text = "Mostrar Codigo Anterior"
        Me.chk_CodigoAnt.UseVisualStyleBackColor = False
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.dtpFechaReporte)
        Me.Frame5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame5.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.Frame5.Location = New System.Drawing.Point(72, 12)
        Me.Frame5.Margin = New System.Windows.Forms.Padding(2)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(113, 27)
        Me.Frame5.TabIndex = 6
        '
        'dtpFechaReporte
        '
        Me.dtpFechaReporte.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFechaReporte.Location = New System.Drawing.Point(4, 5)
        Me.dtpFechaReporte.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFechaReporte.Name = "dtpFechaReporte"
        Me.dtpFechaReporte.Size = New System.Drawing.Size(101, 20)
        Me.dtpFechaReporte.TabIndex = 0
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(65, 47)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(167, 21)
        Me.dbcSucursal.TabIndex = 1
        '
        'dbcCaja
        '
        Me.dbcCaja.Location = New System.Drawing.Point(65, 74)
        Me.dbcCaja.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcCaja.Name = "dbcCaja"
        Me.dbcCaja.Size = New System.Drawing.Size(80, 21)
        Me.dbcCaja.TabIndex = 2
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblVentas.SetIndex(Me._lblVentas_5, CType(5, Integer))
        Me._lblVentas_5.Location = New System.Drawing.Point(23, 17)
        Me._lblVentas_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(46, 20)
        Me._lblVentas_5.TabIndex = 5
        Me._lblVentas_5.Text = "Fecha :"
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblVentas.SetIndex(Me._lblVentas_1, CType(1, Integer))
        Me._lblVentas_1.Location = New System.Drawing.Point(28, 77)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(34, 13)
        Me._lblVentas_1.TabIndex = 8
        Me._lblVentas_1.Text = "Caja :"
        '
        'frmPVDiarioMovtos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(255, 137)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(331, 278)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPVDiarioMovtos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Diario de Movimientos"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame5.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Sub Imprime()
        '        Dim rptDiarioDeMovimientosSecciones As Object
        '        On Error GoTo Merr
        '        Dim aParam(3) As Object
        '        Dim aValues(3) As Object
        '        Dim FechaReporte As Date
        '        Dim cHAVING As String
        '        Dim Encabezado As String
        '        Dim TipoCambioD As Double

        '        If mblnTecleoFechaI Then
        '            Do While (msglTiempoCambioI) <= 2.1
        '            Loop
        '            mblnTecleoFechaI = False
        '        End If
        '        System.Windows.Forms.Application.DoEvents()

        '        '''TipoCambioD = gcurCorpoTIPOCAMBIODOLAR
        '        TipoCambioD = ObtenerTCdelDia()
        '        '''ANTERIOR
        '        '''gStrSql = "SELECT D.Codigo, D.TipoMovto, D.NumPartida, D.Folio, D.Nombre, D.Condicion, D.FolioVenta, D.TipoDEvol, D.Total-D.ImporteVale as Total, " & _
        '        '"D.CodSucursal, ltrim(rtrim(convert(char(20), D.CodArticulo))) as CodArticulo, D.CodigoAnt, D.DescArticulo, D.PrecioReal, D.PrecioLista, Rtrim(Ltrim(C.NombreEmp)) AS NombreEmpresa, A.DescAlmacen AS DescSucursal " & _
        '        '"FROM  DatosParaDiariodeMovimientos_Secciones('" & Format(dtpFechaReporte, C_FORMATFECHAGUARDAR) & "', " & gintCodAlmacen & ", " & gintCodCaja & ", " & TipoCambioD & ") D INNER JOIN  dbo.CatAlmacen A ON D.codsucursal = A.CodAlmacen  CROSS JOIN dbo.ConfiguracionGeneral C  Order by Codigo, Folio, NumPartida "

        '        gStrSql = "SELECT D.Codigo, D.TipoMovto, D.NumPartida, D.Folio, D.Nombre, D.Condicion, D.FolioVenta, D.TipoDEvol, D.Total-D.ImporteVale as Total, D.CodSucursal, ltrim(rtrim(convert(char(20), D.CodArticulo))) as CodArticulo, D.CodigoAnt, D.DescArticulo, " & "Case When TipoMovto = 'V' or TipoMovto = 'A' or TipoMovto = 'XC' or TipoMovto = 'D' Then (D.PrecioReal  * D.Cantidad) Else D.PrecioReal End as PrecioReal, Case When TipoMovto = 'V' or TipoMovto = 'A' or TipoMovto = 'XC' or TipoMovto = 'D' Then (D.PrecioLista * D.Cantidad) Else D.PrecioLista End as PrecioLista, " & "D.Cantidad, Rtrim(Ltrim(C.NombreEmp)) AS NombreEmpresa, A.DescAlmacen AS DescSucursal " & "FROM  DatosParaDiariodeMovimientos_Secciones ('" & Format(dtpFechaReporte.Value, C_FORMATFECHAGUARDAR) & "', " & mintCodSucursal & ", " & mintCodCaja & ", " & TipoCambioD & ") D INNER JOIN  dbo.CatAlmacen A ON D.codsucursal = A.CodAlmacen CROSS JOIN dbo.ConfiguracionGeneral C " & "Order    by Codigo, Folio, NumPartida "
        '        ModEstandar.BorraCmd()
        '        Cmd.CommandText = "dbo.UP_Select_Datos"
        '        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        '        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        '        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        '        frmReportes.rsReport = Cmd.Execute

        '        FechaReporte = dtpFechaReporte.Value
        '        If frmReportes.rsReport.RecordCount = 0 Then
        '            MsgBox("No existe que reportar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
        '            dtpFechaReporte.Focus()
        '        Else
        '            aParam(1) = "FechaReporte"
        '            aParam(2) = "TipoCambio"
        '            aParam(3) = "CodigoArt"
        '            aValues(1) = FechaReporte
        '            aValues(2) = TipoCambioD
        '            aValues(3) = IIf((chk_CodigoAnt.CheckState = System.Windows.Forms.CheckState.Checked), True, False)

        '            frmReportes.Report = rptDiarioDeMovimientosSecciones 'Es el nombre del archivo que se incluyó en el proyecto
        '            frmReportes.Imprime(Me.Text, aParam, aValues)
        '        End If

        'Merr:
        '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub chk_CodigoAnt_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chk_CodigoAnt.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        '   If KeyCode = vbKeyReturn Then dtpFechaReporte.SetFocus
        '   If KeyCode = vbKeyEscape Then dtpFechaReporte.SetFocus
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If FueraChange Then Exit Sub

        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)

        If Trim(dbcSucursal.Text) = "" Then
            mintCodSucursal = 0
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen WHERE TipoAlmacen = 'P' "
        ModDCombo.DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            dtpFechaReporte.Focus()
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(dbcSucursal.Text) & "%'"
        Aux = mintCodSucursal
        mintCodSucursal = 0
        ModDCombo.DCLostFocus(dbcSucursal, gStrSql, mintCodSucursal)
        LlenaDatosCajaSuc((mintCodSucursal))
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        Dim Aux As String
        Aux = Trim(dbcSucursal.Text)
        'If dbcSucursal.SelectedItem <> 0 Then
        'dbcSucursal_Leave(dbcSucursal, New System.EventArgs())
        'End If
        dbcSucursal.Text = Aux
    End Sub

    Private Sub dtpFechareporte_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpFechaReporte.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            eventSender.KeyCode = 0
            Exit Sub
        End If
        '   If KeyCode = vbKeyReturn Then chk_CodigoAnt.SetFocus
    End Sub

    Private Sub dtpFechaReporte_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpFechaReporte.KeyPress
        mblnTecleoFechaI = True
        'msglTiempoCambioI = VB.Timer()
    End Sub

    Private Sub frmPVDiarioMovtos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmPVDiarioMovtos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub Form_Initialize_Renamed()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmPVDiarioMovtos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        ' En este evento del formulario se valida la tecla presionada.
        ' Si es Enter se simula un tab(Avanza al siguiente control)
        ' Si es Escape, se simula un Retroceso de TAB (Regresa al control anterior)
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmPVDiarioMovtos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub Nuevo()
        dtpFechaReporte.Value = Format(Today, C_FORMATFECHAMOSTRAR)
        FueraChange = True
        dbcSucursal.Text = ""
        dbcCaja.Text = ""
        mintCodSucursal = 0
        mintCodCaja = 0
        FueraChange = False
    End Sub

    Private Sub frmPVDiarioMovtos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
        LlenaDatosSucursal((gintCodAlmacen))

    End Sub

    Private Sub frmPVDiarioMovtos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If Not mblnSalir Then
            'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
            ModEstandar.RestaurarForma(Me, False)
            Cancel = 0 'Para que no salga del Formulario hasta que guarde los datos, si no tiene premiso de hacerlo
        Else 'Se quiere salir con escape
            mblnSalir = False
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
                Case MsgBoxResult.Yes 'Sale del Formulario, pero antes preguntar si desea grabar los datos registrados, solo cuando es nuevo
                    Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
                Case MsgBoxResult.No 'No sale del formulario
                    Cancel = 1
                    dtpFechaReporte.Focus()
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmPVDiarioMovtos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub Limpiar()
        Nuevo()
        dtpFechaReporte.Focus()
    End Sub

    Sub LlenaDatosSucursal(ByRef CodSucursal As Integer)
        If CDbl(Numerico(Trim(CStr(CodSucursal)))) = 0 Then Exit Sub
        gStrSql = "SELECT Ltrim(Rtrim(DescAlmacen)) as DescAlmacen From dbo.CatAlmacen Where CodAlmacen =" & CodSucursal & "  And TipoAlmacen = 'P'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            dbcSucursal.Text = RsGral.Fields("DescAlmacen").Value
        Else
            MsgBox("Código de Sucursal no existe." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            CodSucursal = CShort("")
            dbcSucursal.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub dbcCaja_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCaja.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcCaja.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodCaja, RIGHT('00'+LTRIM(CodCaja),2) as NumCaja From dbo.CatCajas Where CodCaja like '" & Trim(dbcCaja.Text) & "%'   And CodAlmacen =" & gintCodAlmacen & " order by NumCAja asc "
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcCaja_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCaja.Enter
        Pon_Tool()
        gStrSql = "SELECT CodCaja, RIGHT('00'+LTRIM(CodCaja),2) as NumCaja From dbo.CatCajas  Where codAlmacen= " & gintCodAlmacen & "  order by NumCAja asc "
        ModDCombo.DCGotFocus(gStrSql, dbcCaja)
    End Sub

    Private Sub dbcCaja_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcCaja.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
    End Sub

    Private Sub dbcCaja_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcCaja.Leave
        mintCodCaja = 0
        gStrSql = "SELECT CodCaja, RIGHT('00'+LTRIM(CodCaja),2) as NumCaja From dbo.CatCajas Where CodCaja like '" & Numerico(dbcCaja.Text) & "%'   and codAlmacen= " & gintCodAlmacen & "    order by NumCAja asc "
        ModDCombo.DCLostFocus(dbcCaja, gStrSql, mintCodCaja)
    End Sub

    Sub LlenaDatosCajaSuc(ByRef CodSucursal As Integer)
        If CDbl(Numerico(Trim(CStr(CodSucursal)))) = 0 Then Exit Sub
        gStrSql = "Select Top 1 CodCaja From CatCajas (Nolock) Where CodAlmacen = " & gintCodAlmacen & " Order by CodCaja "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            FueraChange = True
            dbcCaja.Text = ("00" & Trim(CStr(RsGral.Fields("CodCaja").Value)))
            mintCodCaja = RsGral.Fields("CodCaja").Value
            FueraChange = False
        End If
    End Sub

    Private Function ObtenerTCdelDia() As Decimal
        Dim lSql As String
        Dim Rs As ADODB.Recordset

        'Codigo Utilizado Anteriormente
        '   lSql = "Select FolioVenta, TipoCambio From MovimientosVentasCab(Nolock) " & _
        ''          "Where CodSucursal = " & mintCodSucursal & " And CodCaja = " & mintCodCaja & " And FechaVenta = '" & Format(dtpFechaReporte.Value, C_FORMATFECHAGUARDAR) & "' Order by FolioVenta desc "
        '    ModEstandar.BorraCmd
        '    Cmd.CommandText = "dbo.Up_Select_Datos"
        '    Cmd.CommandType = adCmdStoredProc
        '    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, lSql)
        '    Set Rs = Cmd.Execute
        '    If Rs.RecordCount > 0 Then
        '       Rs.MoveFirst
        '       ObtenerTCdelDia = Rs!TipoCambio
        '    'Else
        '    Else
        '      '''Si no hubo venta busca reparaciones
        '      lSql = "Select FolioReparacion, TipoCambio From Reparaciones (Nolock) " & _
        ''             "Where CodSucursal = " & mintCodSucursal & " And CodCaja = " & mintCodCaja & " And FechaReparacion = '" & Format(dtpFechaReporte.Value, C_FORMATFECHAGUARDAR) & "' Order by FolioReparacion Desc "
        '       ModEstandar.BorraCmd
        '       Cmd.CommandText = "dbo.Up_Select_Datos"
        '       Cmd.CommandType = adCmdStoredProc
        '       Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
        '       Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 800, lSql)
        '       Set Rs = Cmd.Execute
        '       If Rs.RecordCount > 0 Then
        '          Rs.MoveFirst
        '          ObtenerTCdelDia = Rs!TipoCambio
        '       Else
        '          ObtenerTCdelDia = gcurCorpoTIPOCAMBIODOLAR
        '       End If
        '    End If

        '''Nuevo Codigo 25-AGO-2004
        lSql = "SELECT FolioVenta,TipoCambio FROM MovimientosVentasCab(NoLock) " & "WHERE CodSucursal = " & mintCodSucursal & " AND CodCaja = " & mintCodCaja & " AND FechaVenta = '" & Format(dtpFechaReporte.Value, C_FORMATFECHAGUARDAR) & "' ORDER BY FolioVenta DESC "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lSql))
        Rs = Cmd.Execute
        If Rs.RecordCount > 0 Then
            Rs.MoveFirst()
            ObtenerTCdelDia = Rs.Fields("TipoCambio").Value
        Else
            'Si no hubo ventas se busca en los ingresos
            lSql = "SELECT FolioIngreso,TipoCambio FROM Ingresos(NoLock) " & "WHERE CodSucursal = " & mintCodSucursal & " AND CodCaja = " & mintCodCaja & " AND FechaIngreso = '" & Format(dtpFechaReporte.Value, C_FORMATFECHAGUARDAR) & "' ORDER BY FolioIngreso DESC "
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lSql))
            Rs = Cmd.Execute
            If Rs.RecordCount > 0 Then
                Rs.MoveFirst()
                ObtenerTCdelDia = Rs.Fields("TipoCambio").Value
            Else
                'Si no hubo ingresos se busca en las devoluciones
                lSql = "SELECT FolioDevolucion,TipoCambio FROM DevolucionesCab(NoLock) " & "WHERE CodSucursal = " & mintCodSucursal & " AND CodCaja = " & mintCodCaja & " AND FechaDevolucion = '" & Format(dtpFechaReporte.Value, C_FORMATFECHAGUARDAR) & "' ORDER BY FolioDevolucion DESC "
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, lSql))
                Rs = Cmd.Execute
                If Rs.RecordCount > 0 Then
                    Rs.MoveFirst()
                    ObtenerTCdelDia = Rs.Fields("TipoCambio").Value
                Else
                    'Si no hubo devoluciones se toma el tipo de cambio de la configuracion general
                    ObtenerTCdelDia = gcurCorpoTIPOCAMBIODOLAR
                End If
            End If
        End If

    End Function
End Class