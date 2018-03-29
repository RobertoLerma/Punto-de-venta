Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmReportePrestamosPendientes
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    'Programa: Reporte de Apartados
    'Autor: Rosaura Torres López
    'Fecha de Creación:25/Junio/2003
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents dtpFechaReporte As System.Windows.Forms.DateTimePicker
    Public WithEvents Frame5 As System.Windows.Forms.Panel
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents chkTodasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents optPendientes As System.Windows.Forms.RadioButton
    Public WithEvents optRegistrados As System.Windows.Forms.RadioButton
    Public WithEvents fraEstatus As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaIncio As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFin As System.Windows.Forms.DateTimePicker
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents fraPeriodo As System.Windows.Forms.GroupBox
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray

    Dim mblnSalir As Boolean
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim mblnNuevo As Integer

    Sub Imprime()
        Dim rptPrestamosPendientes As New rptPrestamosPendientes
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue


        On Error GoTo Merr
        Dim aParam(3) As Object
        Dim aValues(3) As Object
        Dim FechaInicio As Date
        Dim FechaFin As Date
        Dim cHAVING As String
        Dim Encabezado As String
        If ValidaDatos() = False Then Exit Sub
        FechaInicio = dtpFechaIncio.Value
        FechaFin = dtpFechaFin.Value
        'TextoAdicional = Trim(ModEstandar.QuitaEnter(txtTextoAdicional))
        Encabezado = "Reporte de Préstamos de Artículos Pendientes de Entregar"
        'Validar en que moneda se desea el Reporte
        '    cHAVING = " Having "
        gStrSql = "SELECT     P.FolioAlmacen, P.CodCliente, P.DescCliente, P.Domicilio, P.TelCasa, P.TelOficina, P.FechaAlmacen, P.CodArticulo, P.DescGrupo, P.DescArticulo, " & "P.PrecioVenta, P.NombreEmp, P.Cantidad AS CantidadPrestada, ISNULL(D.Cantidad, 0) AS CantidadDevuelta, P.Cantidad - ISNULL(D.Cantidad, 0) " & "AS Pendiente, P.CodAlmacen, P.DescAlmacen " & "FROM         dbo.PrestamosRealizados('" & VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR) & "', '" & VB6.Format(FechaFin, C_FORMATFECHAGUARDAR) & "', " & C_SalidaPorPrestamodeArticulos & " ) P LEFT OUTER JOIN " & "dbo.DevolucionesdePrestamos('" & VB6.Format(FechaInicio, C_FORMATFECHAGUARDAR) & "', '" & VB6.Format(FechaFin, C_FORMATFECHAGUARDAR) & "', " & C_EntradaPorDevolucionSobrePrestamo & ") D ON P.FolioAlmacen = D.ReferenciaDeOrigen " & "GROUP BY P.FolioAlmacen, P.CodCliente, P.DescCliente, P.Domicilio, P.TelCasa, P.TelOficina, P.FechaAlmacen, P.CodArticulo, P.DescGrupo, P.DescArticulo, " & "P.PrecioVenta , P.NombreEmp, P.Cantidad, IsNull(D.Cantidad, 0), P.Cantidad - IsNull(D.Cantidad, 0), P.CodAlmacen, P.DescAlmacen "

        Select Case True
            Case optPendientes.Checked = True And chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
                cHAVING = " Having  (P.Cantidad - IsNull(D.Cantidad, 0) > 0 ) And " & " CodALmacen =  " & intCodSucursal
            Case optPendientes.Checked = True
                cHAVING = " Having  (P.Cantidad - IsNull(D.Cantidad, 0) > 0 ) "
            Case chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
                cHAVING = "Having  CodALmacen =  " & intCodSucursal
        End Select

        gStrSql = gStrSql & cHAVING

        ModEstandar.BorraCmd()
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute
        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existe que reportar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            Exit Sub
        Else
            rptPrestamosPendientes.SetDataSource(frmReportes.rsReport)
        End If


        'aParam(1) = "FechaInicio"
        'aParam(2) = "FechaFin"
        'aParam(3) = "EncabezadoReporte"
        'aValues(1) = FechaInicio
        'aValues(2) = FechaFin
        'aValues(3) = Encabezado
        'frmReportes.Report = rptPrestamosPendientes 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime("", aParam, aValues) 

        If (FechaInicio <> Nothing) Then
            pdvNum.Value = FechaInicio : pvNum.Add(pdvNum)
            rptPrestamosPendientes.DataDefinition.ParameterFields("FechaInicio").ApplyCurrentValues(pvNum)
        End If

        If (FechaFin <> Nothing) Then
            pdvNum.Value = FechaFin : pvNum.Add(pdvNum)
            rptPrestamosPendientes.DataDefinition.ParameterFields("FechaFin").ApplyCurrentValues(pvNum)
        End If

        If (Encabezado <> Nothing) Then
            pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
            rptPrestamosPendientes.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
        End If

        frmReportes.reporteActual = rptPrestamosPendientes
        frmReportes.Show()

        Cmd.CommandTimeout = 90

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Private Sub chkTodasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodasSucursales.CheckStateChanged
        If chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            dbcSucursales.Text = ""
            intCodSucursal = 0
            dbcSucursales.Enabled = False
        Else
            dbcSucursales.Enabled = True
        End If
    End Sub

    Private Sub chkTodasSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles chkTodasSucursales.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
            KeyCode = 0
        End If
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        mblnNuevo = True
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen  FROM CatAlmacen ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            chkTodasSucursales.Focus()
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        gStrSql = "SELECT CodAlmacen, Ltrim(Rtrim( DescAlmacen )) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        '    LlenaDatos
    End Sub


    Private Sub dtpFechaFin_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpFechaFin.Leave
        'Validar si la FEcha final es Mayor que la Inicial.
        If dtpFechaFin.Value < dtpFechaIncio.Value Then
            MsgBox("La Fecha Final debe ser Mayor que la Inicial." & vbNewLine & "Verifique Por Favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            dtpFechaFin.Value = Today
            dtpFechaIncio.Focus()
            Exit Sub
        End If
        If dtpFechaFin.Value > Today Then
            MsgBox("La fecha final no debe ser mayor que la fecha actual." & vbNewLine & "Verifique por favor...", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            dtpFechaFin.Value = Today
            dtpFechaFin.Focus()
            Exit Sub
        End If
    End Sub

    Function ValidaDatos() As Boolean
        'Validar si la FEcha final es Mayor que la Inicial.
        If dtpFechaFin.Value < dtpFechaIncio.Value Then
            MsgBox("La Fecha Final debe ser Mayor que la Inicial." & vbNewLine & "Verifique Por Favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            dtpFechaFin.Value = Today
            Exit Function
        End If
        If Trim(dbcSucursales.Text) = "" And chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            MsgBox("Proporcione la Sucursal sobre la cual se buscará la Información.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            dbcSucursales.Focus()
            Exit Function
        End If

        ValidaDatos = True
    End Function

    Private Sub dtpFechaIncio_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dtpFechaIncio.KeyDown
        '
        '    ElseIf KeyCode = vbKeyDelete Then
        '        'sI La Tecla presionada fue SUPR, se borrará todo el contenido del form. ya que no es posible hacer modificaciones.
        '        'Unicamnete podran consultarse los datos.
        '        Nuevo
        '   End If
    End Sub

    Private Sub frmReportePrestamosPendientes_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmReportePrestamosPendientes_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub Form_Initialize_Renamed()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmReportePrestamosPendientes_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmReportePrestamosPendientes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub Nuevo()
        dtpFechaIncio.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
        dtpFechaFin.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
        dtpFechaReporte.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
        FueraChange = True
        dbcSucursales.Text = ""
        chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        FueraChange = False
        optRegistrados.Checked = True
    End Sub

    Private Sub frmReportePrestamosPendientes_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        dtpFechaIncio.MinDate = C_FECHAINICIAL
        dtpFechaIncio.MaxDate = C_FECHAFINAL
        dtpFechaFin.MinDate = C_FECHAINICIAL
        dtpFechaFin.MaxDate = C_FECHAFINAL
        Nuevo()
    End Sub

    Private Sub frmReportePrestamosPendientes_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmReportePrestamosPendientes_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub Limpiar()
        Nuevo()
        chkTodasSucursales.Focus()
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.optPendientes = New System.Windows.Forms.RadioButton()
        Me.optRegistrados = New System.Windows.Forms.RadioButton()
        Me.Frame5 = New System.Windows.Forms.Panel()
        Me.dtpFechaReporte = New System.Windows.Forms.DateTimePicker()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me.chkTodasSucursales = New System.Windows.Forms.CheckBox()
        Me.fraEstatus = New System.Windows.Forms.GroupBox()
        Me.fraPeriodo = New System.Windows.Forms.GroupBox()
        Me.dtpFechaIncio = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFin = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame5.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.fraEstatus.SuspendLayout()
        Me.fraPeriodo.SuspendLayout()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.Color.Black
        Me._Label1_1.Location = New System.Drawing.Point(16, 44)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(60, 17)
        Me._Label1_1.TabIndex = 4
        Me._Label1_1.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_1, "Nombre de la Farmacia Actual")
        '
        'optPendientes
        '
        Me.optPendientes.BackColor = System.Drawing.SystemColors.Control
        Me.optPendientes.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPendientes.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.optPendientes.Location = New System.Drawing.Point(174, 20)
        Me.optPendientes.Name = "optPendientes"
        Me.optPendientes.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPendientes.Size = New System.Drawing.Size(159, 17)
        Me.optPendientes.TabIndex = 14
        Me.optPendientes.TabStop = True
        Me.optPendientes.Text = "Solo préstamos pendientes"
        Me.ToolTip1.SetToolTip(Me.optPendientes, "Prétamos Pendietes")
        Me.optPendientes.UseVisualStyleBackColor = False
        '
        'optRegistrados
        '
        Me.optRegistrados.BackColor = System.Drawing.SystemColors.Control
        Me.optRegistrados.Checked = True
        Me.optRegistrados.Cursor = System.Windows.Forms.Cursors.Default
        Me.optRegistrados.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.optRegistrados.Location = New System.Drawing.Point(27, 20)
        Me.optRegistrados.Name = "optRegistrados"
        Me.optRegistrados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optRegistrados.Size = New System.Drawing.Size(130, 17)
        Me.optRegistrados.TabIndex = 12
        Me.optRegistrados.TabStop = True
        Me.optRegistrados.Text = "Todos los préstamos"
        Me.ToolTip1.SetToolTip(Me.optRegistrados, "Préstamos Registrados")
        Me.optRegistrados.UseVisualStyleBackColor = False
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.Controls.Add(Me.dtpFechaReporte)
        Me.Frame5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame5.Enabled = False
        Me.Frame5.ForeColor = System.Drawing.SystemColors.InactiveBorder
        Me.Frame5.Location = New System.Drawing.Point(240, 0)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(111, 33)
        Me.Frame5.TabIndex = 13
        '
        'dtpFechaReporte
        '
        Me.dtpFechaReporte.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaReporte.Location = New System.Drawing.Point(8, 8)
        Me.dtpFechaReporte.Name = "dtpFechaReporte"
        Me.dtpFechaReporte.Size = New System.Drawing.Size(100, 20)
        Me.dtpFechaReporte.TabIndex = 1
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.dbcSucursales)
        Me.Frame1.Controls.Add(Me.chkTodasSucursales)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 32)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(345, 73)
        Me.Frame1.TabIndex = 2
        Me.Frame1.TabStop = False
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(74, 40)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(259, 21)
        Me.dbcSucursales.TabIndex = 5
        '
        'chkTodasSucursales
        '
        Me.chkTodasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodasSucursales.Checked = True
        Me.chkTodasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodasSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodasSucursales.Location = New System.Drawing.Point(16, 16)
        Me.chkTodasSucursales.Name = "chkTodasSucursales"
        Me.chkTodasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodasSucursales.Size = New System.Drawing.Size(153, 18)
        Me.chkTodasSucursales.TabIndex = 3
        Me.chkTodasSucursales.Text = "Todas las Sucursales"
        Me.chkTodasSucursales.UseVisualStyleBackColor = False
        '
        'fraEstatus
        '
        Me.fraEstatus.BackColor = System.Drawing.SystemColors.Control
        Me.fraEstatus.Controls.Add(Me.optPendientes)
        Me.fraEstatus.Controls.Add(Me.optRegistrados)
        Me.fraEstatus.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraEstatus.Location = New System.Drawing.Point(8, 176)
        Me.fraEstatus.Name = "fraEstatus"
        Me.fraEstatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraEstatus.Size = New System.Drawing.Size(345, 49)
        Me.fraEstatus.TabIndex = 11
        Me.fraEstatus.TabStop = False
        Me.fraEstatus.Text = " Seleccionar..."
        '
        'fraPeriodo
        '
        Me.fraPeriodo.BackColor = System.Drawing.SystemColors.Control
        Me.fraPeriodo.Controls.Add(Me.dtpFechaIncio)
        Me.fraPeriodo.Controls.Add(Me.dtpFechaFin)
        Me.fraPeriodo.Controls.Add(Me.Label2)
        Me.fraPeriodo.Controls.Add(Me._Label1_0)
        Me.fraPeriodo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraPeriodo.Location = New System.Drawing.Point(8, 112)
        Me.fraPeriodo.Name = "fraPeriodo"
        Me.fraPeriodo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraPeriodo.Size = New System.Drawing.Size(345, 57)
        Me.fraPeriodo.TabIndex = 6
        Me.fraPeriodo.TabStop = False
        Me.fraPeriodo.Text = " Período "
        '
        'dtpFechaIncio
        '
        Me.dtpFechaIncio.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaIncio.Location = New System.Drawing.Point(64, 21)
        Me.dtpFechaIncio.Name = "dtpFechaIncio"
        Me.dtpFechaIncio.Size = New System.Drawing.Size(105, 20)
        Me.dtpFechaIncio.TabIndex = 8
        '
        'dtpFechaFin
        '
        Me.dtpFechaFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFin.Location = New System.Drawing.Point(232, 21)
        Me.dtpFechaFin.Name = "dtpFechaFin"
        Me.dtpFechaFin.Size = New System.Drawing.Size(105, 20)
        Me.dtpFechaFin.TabIndex = 10
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(202, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(33, 21)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Al :"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Location = New System.Drawing.Point(29, 27)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(29, 21)
        Me._Label1_0.TabIndex = 7
        Me._Label1_0.Text = "Del :"
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me._lblVentas_5.Location = New System.Drawing.Point(199, 13)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(43, 12)
        Me._lblVentas_5.TabIndex = 0
        Me._lblVentas_5.Text = "Fecha :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(123, 243)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 97
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(8, 243)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 96
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmReportePrestamosPendientes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(360, 288)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.fraEstatus)
        Me.Controls.Add(Me.fraPeriodo)
        Me.Controls.Add(Me._lblVentas_5)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(389, 251)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReportePrestamosPendientes"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de préstamos de artículos"
        Me.Frame5.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.fraEstatus.ResumeLayout(False)
        Me.fraPeriodo.ResumeLayout(False)
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class