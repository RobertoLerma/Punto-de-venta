Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmRptTransferenciasNoConciliadas
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    'Programa: Reporte de Transferencias no conciliadas
    'Autor: Rosaura Torres López
    'Fecha de Creación: 22/Septiembe/2003
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkDetallarporTransferencia As System.Windows.Forms.CheckBox
    Public WithEvents chkNoConciliados As System.Windows.Forms.CheckBox
    Public WithEvents chkTodaslasSucursales As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucDestino As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_2 As System.Windows.Forms.Label
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents dtpFechaInicio As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpFechaFin As System.Windows.Forms.DateTimePicker
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents fraPeriodo As System.Windows.Forms.GroupBox
    Public WithEvents dbcSucOrigen As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Frame1_0 As System.Windows.Forms.GroupBox
    Public WithEvents Frame1 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Dim mblnSalir As Boolean
    'Dim mintJFamilia As Integer
    'Dim mintJLinea As Integer
    'Dim mintJSubLinea As Integer
    'Dim mintVFamilia As Integer
    'Dim mintVLinea As Integer
    'Dim mintRMarca As Integer
    'Dim mintRModelo As Integer
    'Dim intCodSucOrigen As Integer
    'Dim IntCodOrigen As Integer
    Dim mblnFueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucOrigen As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim intCodSucDestino As Integer

    Sub Imprime()
        Dim rptTransferenciasnoConciliadas As New rptTransferenciasnoConciliadas
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr
        Dim aParam(5) As Object
        Dim aValues(5) As Object
        'Dim FechaInicio As Date
        'Dim FechaFin As Date
        Dim TextoAdicional As String
        Dim Encabezado As String
        Dim ConsultaGuardar As String
        Dim ConsultaReporte As String
        Dim mblnTRansaccion As Boolean
        If ValidaDatos() = False Then Exit Sub

        Encabezado = "Reporte de transferencias no conciliadas"

        Dim fechaInicio As String = AgregarHoraAFecha(dtpFechaInicio.Value)
        Dim fechaFin As String = AgregarHoraAFecha(dtpFechaFin.Value)

        gStrSql = "SELECT MC.FolioAlmacen AS FolioEntrada, MC.ReferenciaDeOrigen AS FolioSalida, MC.CodAlmacen AS CodAlmacenEntrada, LTRIM(RTRIM(Al.DescAlmacen)) 
" & "AS DescAlmacenEntrada, MC.CodAlmacenRef AS CodAlmacenSalida, LTRIM(RTRIM(Al_1.DescAlmacen)) AS DEscAlmacenSalida, MD.CodArticulo, 
" & "dbo.CatArticulos.DescArticulo, U.DescUnidad, MD.Cantidad, MD.Confirmacion, LTRIM(RTRIM(CG.NombreEmp)) AS NombreEmpresa ,TC.ConciliadoTotal 
" & "FROM dbo.MovtosAlmacenCab MC INNER JOIN " & "dbo.MovtosAlmacenDet MD ON MC.FolioAlmacen = MD.FolioAlmacen INNER JOIN 
" & "dbo.CatArticulos ON MD.CodArticulo = dbo.CatArticulos.CodArticulo INNER JOIN " & "dbo.CatAlmacen Al ON MC.CodAlmacen = Al.CodAlmacen INNER JOIN 
" & "dbo.CatAlmacen Al_1 ON MC.CodAlmacenRef = Al_1.CodAlmacen INNER JOIN " & "dbo.CatUnidades U ON dbo.CatArticulos.CodUnidad = U.CodUnidad CROSS JOIN 
" & "dbo.ConfiguracionGeneral CG  INNER JOIN 
" & "dbo.TransferenciasConciliadas(" & C_EntradaPorTransferencia & "," & intCodSucOrigen & ",'" & fechaInicio & "', '" & fechaFin & "') TC ON MC.FolioAlmacen = TC.FolioEntrada 
" & "Where (MC.CodMovtoAlm = " & C_EntradaPorTransferencia & ") And (MC.CodALmacenREf =" & intCodSucOrigen & " )  
" & "And Mc.FechaAlmacen BetWeen '" & fechaInicio & "'  and '" & fechaFin & "'  "

        If chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            gStrSql = gStrSql & " And (mc.CodAlmacen = " & intCodSucDestino & " )"
        End If
        If chkNoConciliados.CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = gStrSql & " And (ConciliadoTotal = 0 )"
        End If


        ModEstandar.BorraCmd()
        'Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existe que reportar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            Exit Sub
        Else
            rptTransferenciasnoConciliadas.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "EncabezadoReporte"
        'aValues(1) = Encabezado
        'aParam(2) = "FechaInicio"
        'aValues(2) = dtpFechaInicio.Value
        'aParam(3) = "FechaFin"
        'aValues(3) = dtpFechaFin.Value
        'aParam(4) = "MostrarDetalle"
        'aValues(4) = IIf(chkDetallarporTransferencia.CheckState = System.Windows.Forms.CheckState.Checked, True, False)
        'frmReportes.Report = rptTransferenciasnoConciliadas 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime(Me.Text, aParam, aValues)

        If (Encabezado <> Nothing) Then
            pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
            rptTransferenciasnoConciliadas.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
        End If

        If (dtpFechaInicio.Value <> Nothing) Then
            pdvNum.Value = dtpFechaInicio.Value : pvNum.Add(pdvNum)
            rptTransferenciasnoConciliadas.DataDefinition.ParameterFields("FechaInicio").ApplyCurrentValues(pvNum)
        End If

        If (dtpFechaFin.Value <> Nothing) Then
            pdvNum.Value = dtpFechaFin.Value : pvNum.Add(pdvNum)
            rptTransferenciasnoConciliadas.DataDefinition.ParameterFields("FechaFin").ApplyCurrentValues(pvNum)
        End If
        If (chkDetallarporTransferencia.CheckState = System.Windows.Forms.CheckState.Checked Or chkDetallarporTransferencia.CheckState = System.Windows.Forms.CheckState.Unchecked <> Nothing) Then
            pdvNum.Value = IIf(chkDetallarporTransferencia.CheckState = System.Windows.Forms.CheckState.Checked, True, False) : pvNum.Add(pdvNum)
            rptTransferenciasnoConciliadas.DataDefinition.ParameterFields("MostrarDetalle").ApplyCurrentValues(pvNum)
        End If

        frmReportes.reporteActual = rptTransferenciasnoConciliadas
        frmReportes.Show()

        'Cmd.CommandTimeout = 90
        Exit Sub

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub
    Function ValidaDatos() As Boolean
        If Trim(dbcSucDestino.Text) = "" And chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            MsgBox("Proporcione la sucursal destino.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            dbcSucDestino.Focus()
            Exit Function
        End If
        If Trim(dbcSucOrigen.Text) = "" Then
            MsgBox("Proporcione la sucursal origen.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            dbcSucOrigen.Focus()
            Exit Function
        End If
        If CDate(dtpFechaInicio.Value) > Today Then
            MsgBox("La fecha inicial debe ser menor o igual a la de hoy." & vbNewLine & "Verifique por favor.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            dtpFechaInicio.Focus()
            Exit Function
        End If
        If CDate(dtpFechaFin.Value) > Today Then
            MsgBox("La fecha final debe ser menor o igual a la de hoy." & vbNewLine & "Verifique por favor.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            dtpFechaFin.Focus()
            Exit Function
        End If
        If dtpFechaFin.Value <dtpFechaInicio.Value Then
            MsgBox("La fecha final debe ser mayor o igual a la inicial." & vbNewLine & "Verifique por favor.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            dtpFechaFin.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub chkTodaslasSucursales_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodaslasSucursales.CheckStateChanged
        mblnFueraChange = True
        If chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            dbcSucDestino.Text = ""
            dbcSucDestino.Enabled = False
        Else
            dbcSucDestino.Text = ""
            dbcSucDestino.Enabled = True
        End If
        mblnFueraChange = False
    End Sub

    Private Sub dbcsucdestino_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucDestino.CursorChanged

        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcsucDestino" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucDestino.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucOrigen = 0
        mblnFueraChange = True

        mblnFueraChange = False
    End Sub

    Private Sub dbcsucdestino_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucDestino.Enter
        '    If Screen.ActiveForm.ActiveControl.Name <> dbcsucDestino.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen where TipoAlmacen ='P'ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucDestino)
    End Sub

    Private Sub dbcsucdestino_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucDestino.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucDestino_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucDestino.KeyUp
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Up Or eventArgs.KeyCode = System.Windows.Forms.Keys.Down Then
            PonerCodigoSucursal(dbcSucDestino)
            Exit Sub
        End If
    End Sub

    Private Sub dbcsucdestino_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucDestino.Leave
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucDestino.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucDestino, gStrSql, intCodSucDestino)
    End Sub

    Private Sub dbcSucDestino_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSucDestino.MouseUp
        PonerCodigoSucursal(dbcSucDestino)
    End Sub

    Private Sub dbcsucorigen_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucOrigen.CursorChanged

        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcsucDestino" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,LTRIM(RTIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucOrigen.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucOrigen = 0
        mblnFueraChange = True

        mblnFueraChange = False
    End Sub

    Private Sub dbcsucorigen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucOrigen.Enter
        '    If Screen.ActiveForm.ActiveControl.Name <> dbcsucDestino.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen where TipoAlmacen ='P'ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucOrigen)
    End Sub

    Private Sub dbcsucorigen_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucOrigen.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucOrigen_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucOrigen.KeyUp
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Up Or eventArgs.KeyCode = System.Windows.Forms.Keys.Down Then
            PonerCodigoSucursal(dbcSucOrigen)
            Exit Sub
        End If
    End Sub

    Private Sub dbcsucorigen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucOrigen.Leave
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucOrigen.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucOrigen, gStrSql, intCodSucOrigen)
        btnImprimir.Focus()
    End Sub

    Private Sub dbcSucOrigen_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSucOrigen.MouseUp
        'PonerCodigoSucursal(dbcSucOrigen)
    End Sub

    'Private Sub dbcsucdestino_Change()
    '
    '    If mblnFueraChange = True Then Exit Sub
    '    If Screen.ActiveForm.ActiveControl.Name <> "dbcsucDestino" Then
    '        Exit Sub
    '    End If
    '    gStrSql = "SELECT CodAlmacen,LTRIM(RTIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucDestino) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
    '    DCChange gStrSql, Tecla
    '    intCodSucDestino = 0
    '    mblnFueraChange = True
    '    txtCodSucursal = ""
    '    mblnFueraChange = False
    'End Sub
    '
    'Private Sub dbcsucdestino_GotFocus()
    ''    If Screen.ActiveForm.ActiveControl.Name <> dbcsucDestino.Name Then Exit Sub
    '    Pon_Tool
    '    gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen where TipoAlmacen ='P'ORDER BY DescAlmacen"
    '    DCGotFocus gStrSql
    'End Sub
    '
    'Private Sub dbcsucdestino_KeyDown(KeyCode As Integer, Shift As Integer)
    '    If KeyCode = vbKeyEscape Then
    '        ModEstandar.RetrocederTab Me
    '    End If
    '    Tecla = KeyCode
    'End Sub
    '
    'Private Sub dbcsucdestino_LostFocus()
    '    gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucDestino) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
    '    DCLostFocus dbcSucDestino, gStrSql, intCodSucDestino
    'End Sub

    Private Sub frmRptTransferenciasNoConciliadas_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmRptTransferenciasNoConciliadas_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub
    Private Sub Form_Initialize_Renamed()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmRptTransferenciasNoConciliadas_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmRptTransferenciasNoConciliadas_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub Nuevo()

        intCodSucOrigen = 0
        intCodSucDestino = 0
        chkNoConciliados.CheckState = System.Windows.Forms.CheckState.Checked
        chkDetallarporTransferencia.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodaslasSucursales.CheckState = System.Windows.Forms.CheckState.Checked
        mblnFueraChange = True
        dtpFechaInicio.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
        dtpFechaFin.Value = VB6.Format(Today, C_FORMATFECHAMOSTRAR)
        dbcSucDestino.Text = ""
        dbcSucOrigen.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub frmRptTransferenciasNoConciliadas_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        dtpFechaInicio.MinDate = C_FECHAINICIAL
        dtpFechaInicio.MaxDate = C_FECHAFINAL
        dtpFechaFin.MinDate = C_FECHAINICIAL
        dtpFechaFin.MaxDate = C_FECHAFINAL
        Nuevo()
        '    txtCodSucursal = gintCodAlmacen
        '    txtCodSucursal_LostFocus
    End Sub

    Private Sub frmRptTransferenciasNoConciliadas_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If Not mblnSalir Then
        '    'Si se desea cerrar la forma y esta se encuentra minimizada, ésta se restaura
        '    ModEstandar.RestaurarForma(Me, False)
        '    Cancel = 0 'Para que no salga del Formulario hasta que guarde los datos, si no tiene premiso de hacerlo
        'Else 'Se quiere salir con escape
        '    mblnSalir = False
        '    Select Case MsgBox(C_msgSALIR, MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrCorpoNOMBREEMPRESA)
        '        Case MsgBoxResult.Yes 'Sale del Formulario, pero antes preguntar si desea grabar los datos registrados, solo cuando es nuevo
        '            Cancel = 0 'Sale de la Captura, Con 1: Sigue en la captura
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmRptTransferenciasNoConciliadas_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub Limpiar()
        Nuevo()
        dbcSucOrigen.Focus()
    End Sub

    Sub PonerCodigoSucursal(ByRef ControlLlamado As System.Windows.Forms.ComboBox)
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(ControlLlamado.Text) & "' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'DCLostFocus dbcAlmacenSalida, gStrSql, intCodSucursal
        If RsGral.RecordCount > 0 Then
            mblnFueraChange = True
            If ControlLlamado Is dbcSucDestino Then
                intCodSucDestino = RsGral.Fields("CodAlmacen").Value
            ElseIf ControlLlamado Is dbcSucOrigen Then
                intCodSucOrigen = RsGral.Fields("CodAlmacen").Value
            End If
            mblnFueraChange = False
        End If
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me._Frame1_0 = New System.Windows.Forms.GroupBox()
        Me.chkDetallarporTransferencia = New System.Windows.Forms.CheckBox()
        Me.chkNoConciliados = New System.Windows.Forms.CheckBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.chkTodaslasSucursales = New System.Windows.Forms.CheckBox()
        Me.dbcSucDestino = New System.Windows.Forms.ComboBox()
        Me._Label1_2 = New System.Windows.Forms.Label()
        Me.fraPeriodo = New System.Windows.Forms.GroupBox()
        Me.dtpFechaInicio = New System.Windows.Forms.DateTimePicker()
        Me.dtpFechaFin = New System.Windows.Forms.DateTimePicker()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dbcSucOrigen = New System.Windows.Forms.ComboBox()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.Frame1 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me._Frame1_0.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.fraPeriodo.SuspendLayout()
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_Frame1_0
        '
        Me._Frame1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_0.Controls.Add(Me.chkDetallarporTransferencia)
        Me._Frame1_0.Controls.Add(Me.chkNoConciliados)
        Me._Frame1_0.Controls.Add(Me.Frame2)
        Me._Frame1_0.Controls.Add(Me.fraPeriodo)
        Me._Frame1_0.Controls.Add(Me.dbcSucOrigen)
        Me._Frame1_0.Controls.Add(Me._Label1_1)
        Me._Frame1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.SetIndex(Me._Frame1_0, CType(0, Short))
        Me._Frame1_0.Location = New System.Drawing.Point(8, 0)
        Me._Frame1_0.Name = "_Frame1_0"
        Me._Frame1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_0.Size = New System.Drawing.Size(457, 257)
        Me._Frame1_0.TabIndex = 12
        Me._Frame1_0.TabStop = False
        '
        'chkDetallarporTransferencia
        '
        Me.chkDetallarporTransferencia.BackColor = System.Drawing.SystemColors.Control
        Me.chkDetallarporTransferencia.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDetallarporTransferencia.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDetallarporTransferencia.Location = New System.Drawing.Point(8, 232)
        Me.chkDetallarporTransferencia.Name = "chkDetallarporTransferencia"
        Me.chkDetallarporTransferencia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDetallarporTransferencia.Size = New System.Drawing.Size(177, 17)
        Me.chkDetallarporTransferencia.TabIndex = 11
        Me.chkDetallarporTransferencia.Text = "Detallar por transferencia"
        Me.chkDetallarporTransferencia.UseVisualStyleBackColor = False
        '
        'chkNoConciliados
        '
        Me.chkNoConciliados.BackColor = System.Drawing.SystemColors.Control
        Me.chkNoConciliados.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkNoConciliados.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkNoConciliados.Location = New System.Drawing.Point(8, 208)
        Me.chkNoConciliados.Name = "chkNoConciliados"
        Me.chkNoConciliados.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkNoConciliados.Size = New System.Drawing.Size(129, 17)
        Me.chkNoConciliados.TabIndex = 10
        Me.chkNoConciliados.Text = "No conciliados"
        Me.chkNoConciliados.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.chkTodaslasSucursales)
        Me.Frame2.Controls.Add(Me.dbcSucDestino)
        Me.Frame2.Controls.Add(Me._Label1_2)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(8, 56)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(433, 81)
        Me.Frame2.TabIndex = 2
        Me.Frame2.TabStop = False
        Me.Frame2.Text = " Almacén destino "
        '
        'chkTodaslasSucursales
        '
        Me.chkTodaslasSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodaslasSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodaslasSucursales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodaslasSucursales.Location = New System.Drawing.Point(16, 24)
        Me.chkTodaslasSucursales.Name = "chkTodaslasSucursales"
        Me.chkTodaslasSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodaslasSucursales.Size = New System.Drawing.Size(129, 17)
        Me.chkTodaslasSucursales.TabIndex = 3
        Me.chkTodaslasSucursales.Text = "Todas las Sucursales"
        Me.chkTodaslasSucursales.UseVisualStyleBackColor = False
        '
        'dbcSucDestino
        '
        Me.dbcSucDestino.Location = New System.Drawing.Point(128, 48)
        Me.dbcSucDestino.Name = "dbcSucDestino"
        Me.dbcSucDestino.Size = New System.Drawing.Size(299, 21)
        Me.dbcSucDestino.TabIndex = 4
        '
        '_Label1_2
        '
        Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_2, CType(2, Short))
        Me._Label1_2.Location = New System.Drawing.Point(64, 48)
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.Size = New System.Drawing.Size(73, 17)
        Me._Label1_2.TabIndex = 13
        Me._Label1_2.Text = "Sucursal :"
        '
        'fraPeriodo
        '
        Me.fraPeriodo.BackColor = System.Drawing.SystemColors.Control
        Me.fraPeriodo.Controls.Add(Me.dtpFechaInicio)
        Me.fraPeriodo.Controls.Add(Me.dtpFechaFin)
        Me.fraPeriodo.Controls.Add(Me._Label1_0)
        Me.fraPeriodo.Controls.Add(Me.Label3)
        Me.fraPeriodo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraPeriodo.Location = New System.Drawing.Point(8, 144)
        Me.fraPeriodo.Name = "fraPeriodo"
        Me.fraPeriodo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraPeriodo.Size = New System.Drawing.Size(433, 57)
        Me.fraPeriodo.TabIndex = 5
        Me.fraPeriodo.TabStop = False
        Me.fraPeriodo.Text = " Período "
        '
        'dtpFechaInicio
        '
        Me.dtpFechaInicio.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaInicio.Location = New System.Drawing.Point(72, 20)
        Me.dtpFechaInicio.Name = "dtpFechaInicio"
        Me.dtpFechaInicio.Size = New System.Drawing.Size(105, 20)
        Me.dtpFechaInicio.TabIndex = 7
        '
        'dtpFechaFin
        '
        Me.dtpFechaFin.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaFin.Location = New System.Drawing.Point(280, 20)
        Me.dtpFechaFin.Name = "dtpFechaFin"
        Me.dtpFechaFin.Size = New System.Drawing.Size(105, 20)
        Me.dtpFechaFin.TabIndex = 9
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(32, 23)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(25, 21)
        Me._Label1_0.TabIndex = 6
        Me._Label1_0.Text = "Del :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(248, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(33, 21)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Al :"
        '
        'dbcSucOrigen
        '
        Me.dbcSucOrigen.Location = New System.Drawing.Point(136, 24)
        Me.dbcSucOrigen.Name = "dbcSucOrigen"
        Me.dbcSucOrigen.Size = New System.Drawing.Size(299, 21)
        Me.dbcSucOrigen.TabIndex = 1
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(16, 24)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(113, 17)
        Me._Label1_1.TabIndex = 0
        Me._Label1_1.Text = "Almacén origen :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 280)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 101
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(12, 280)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 100
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmRptTransferenciasNoConciliadas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(469, 328)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me._Frame1_0)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRptTransferenciasNoConciliadas"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de Tranferencias no conciliadas"
        Me._Frame1_0.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.fraPeriodo.ResumeLayout(False)
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class