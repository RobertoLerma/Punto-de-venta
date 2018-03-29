Option Strict Off
Option Explicit On
Imports System.IO
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility

Public Class frmVtasRPTVentasSalidadeMercanciaComisionVendedor
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkTodosVendedores As System.Windows.Forms.CheckBox
    Public WithEvents chkTodas As System.Windows.Forms.CheckBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_3 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_3 As System.Windows.Forms.GroupBox
    Public WithEvents dbcVendedor As System.Windows.Forms.ComboBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean

    Dim mblnFueraChange As Boolean
    Dim mintCodSucursal As Integer
    Dim mintCodVendedor As Integer
    Dim tecla As Integer

    Dim cTablaTmp As String
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Dim mblnSalir As Boolean

    Public Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.chkTodas.Focus()
    End Sub

    Public Sub Nuevo()
        Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodas_CheckStateChanged(chkTodas, New System.EventArgs())

        Me.chkTodosVendedores.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodosVendedores_CheckStateChanged(chkTodosVendedores, New System.EventArgs())

        Me.dtpDesde.Value = Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = Format(Today, "dd/MMM/yyyy")

        Me.txtMensaje.Text = ""

        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
    End Sub

    Function DevuelveQuery() As String
        On Error GoTo Err_Renamed
        Dim Sql As String

        '''Sql = "SELECT ING.CodSucursal,CAST(CA.DescAlmacen AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS DescAlmacen,ING.CodVendedor,CAST(CV.DescVendedor AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS DescVendedor,CAST(Ing.FolioIngreso AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioIngreso,CAST(VTA.FolioVenta AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioVenta,Ing.FechaIngreso,'Venta' AS Movimiento," & _
        '"ROUND(SUM(PrecioReal * (Cantidad - CantidadDev) + CASE WHEN NumPartida = 1 THEN Redondeo ELSE 0 END),2) AS Importe," & _
        '"ING.Total,ING.Total * (CCV.PorcComision / 100) AS Comision " & _
        '"FROM DBO.VTAS_SALIDAMCIA('01/01/1900','" & Format(Date, C_FORMATFECHAGUARDAR) & "') VTA " & _
        '"INNER JOIN Ingresos Ing ON CAST(VTA.FolioVenta AS NVarChar) COLLATE Traditional_Spanish_CI_AI = CAST(Ing.FolioMovto AS NVarChar) COLLATE Traditional_Spanish_CI_AI INNER JOIN (SELECT * FROM CatAlmacen WHERE TipoAlmacen = 'P') CA ON ING.CodSucursal = CA.CodAlmacen INNER JOIN CatVendedores CV ON ING.CodVendedor = CV.CodVendedor INNER JOIN CatComisionXVendedor CCV ON MONTH(ING.FechaIngreso) = MONTH(CCV.FechaPeriodo) " & _
        '"WHERE (Cantidad - CantidadDev) > 0 AND Ing.FechaIngreso BETWEEN '" & Format(dtpDesde, C_FORMATFECHAGUARDAR) & "' AND '" & Format(dtpHasta, C_FORMATFECHAGUARDAR) & "' " & _
        'IIf(mintCodSucursal <> 0, "AND ING.CodSucursal = " & mintCodSucursal & " ", "") & IIf(mintCodVendedor <> 0, "AND ING.CodVendedor = " & mintCodVendedor & " ", "") & _
        '"GROUP BY ING.CodSucursal,CA.DescAlmacen,ING.CodVendedor,CV.DescVendedor,Ing.FolioIngreso,VTA.FolioVenta , ING.FechaIngreso, ING.Total, CCV.PorcComision " & _
        '"UNION " & _
        '"SELECT ING.CodSucursal,CAST(CA.DescAlmacen AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS DescAlmacen,ING.CodVendedor,CAST(CV.DescVendedor AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS DescVendedor,CAST(ING.FolioIngreso AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioIngreso,CAST(ING.FolioMovto AS NVarChar) COLLATE Traditional_Spanish_CI_AI AS FolioMovto,Ing.FechaIngreso,'Devolucion' AS Movimiento," & _
        '"ING.Total AS Importe,0 as Total,ING.Total * (CCV.PorcComision / 100) AS Comision " & _
        '"FROM Ingresos Ing INNER JOIN (SELECT * FROM CatAlmacen WHERE TipoAlmacen = 'P') CA ON ING.CodSucursal = CA.CodAlmacen INNER JOIN CatVendedores CV ON ING.CodVendedor = CV.CodVendedor INNER JOIN CatComisionXVendedor CCV ON MONTH(ING.FechaIngreso) = MONTH(CCV.FechaPeriodo) " & _
        '"WHERE Ing.FechaIngreso BETWEEN '" & Format(dtpDesde, C_FORMATFECHAGUARDAR) & "' AND '" & Format(dtpHasta, C_FORMATFECHAGUARDAR) & "' AND ING.TipoIngreso = 'D' " & _
        'IIf(mintCodSucursal <> 0, "AND ING.CodSucursal = " & mintCodSucursal & " ", "") & IIf(mintCodVendedor <> 0, "AND ING.CodVendedor = " & mintCodVendedor & " ", "")

        Sql = ""
        Sql = Sql & "Select   Ing.CodSucursal, CA.DescAlmacen, Ing.CodVendedor, CV.DescVendedor, Ing.FolioIngreso, Ing.FolioMovto as FolioVenta, Ing.FechaIngreso, Movimiento, Ing.Total, (Ing.TotalIngreso-Ing.ImporteVale) as TotalIngreso, ((Ing.TotalIngreso-Ing.ImporteVale) * (CxV.PorcComision / 100)) AS Comision " & vbNewLine
        Sql = Sql & "From     ( " & vbNewLine
        Sql = Sql & "         Select   A.FolioIngreso, A.FechaIngreso, A.CodSucursal, A.FolioMovto, A.CodVendedor, C.Total+C.Redondeo as TOTAL, A.Total as TotalIngreso, B.TipoCambio, 'Venta' AS Movimiento, Sum(Case When FP.EsDevolucion = 1 Then Case When FP.EsDolar = 1 then Round(B.Importe,4) Else Round(B.Importe/A.TipoCambio,4) End Else 0 End ) as ImporteVale " & vbNewLine
        Sql = Sql & "         From     Ingresos A (Nolock) " & vbNewLine
        Sql = Sql & "                  Inner Join IngresosFormadePago B On A.FolioIngreso = B.FolioIngreso " & vbNewLine
        Sql = Sql & "                  Inner Join CatFormasPago FP (Nolock) On B.CodFormaPago = FP.CodFormaPago " & vbNewLine
        Sql = Sql & "                  Inner Join MovimientosVentasCab C (Nolock) On A.FolioMovto = C.FolioVenta " & vbNewLine
        Sql = Sql & "         Where    A.FechaIngreso between '" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "' And '" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "' " & vbNewLine

        Sql = Sql & IIf(mintCodSucursal <> 0, "         And      A.CodSucursal = " & mintCodSucursal & " ", " ") & vbNewLine
        Sql = Sql & IIf(mintCodVendedor <> 0, "         And      A.CodVendedor = " & mintCodVendedor & " ", " ") & vbNewLine

        ''' 06AGO2007 - MAVF
        Sql = Sql & "         Group    by A.FolioIngreso, A.FechaIngreso, A.CodSucursal, A.FolioMovto, A.CodVendedor, C.Moneda, C.Total+C.Redondeo, A.Total, B.TipoCambio " & vbNewLine
        Sql = Sql & "         ) Ing " & vbNewLine
        Sql = Sql & "         INNER JOIN ( " & vbNewLine
        Sql = Sql & "         SELECT   * FROM CatAlmacen (Nolock) WHERE TipoAlmacen = 'P' " & vbNewLine
        Sql = Sql & "         ) CA ON Ing.CodSucursal = CA.CodAlmacen " & vbNewLine
        Sql = Sql & "         INNER JOIN CatVendedores CV (Nolock) ON Ing.CodVendedor = CV.CodVendedor " & vbNewLine
        Sql = Sql & "         INNER JOIN CatComisionXVendedor CXV (Nolock) ON MONTH(Ing.FechaIngreso) = MONTH(CxV.FechaPeriodo) AND YEAR(Ing.FechaIngreso) = YEAR(CxV.FechaPeriodo) " & vbNewLine
        Sql = Sql & "Order    by Ing.CodSucursal, Ing.CodVendedor, Ing.FechaIngreso, Ing.FolioIngreso "

        DevuelveQuery = Sql
Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    Public Sub Imprime()
        Dim rptVentasSalidaDeMercanciaComisionVendedor As New rptVentasSalidaDeMercanciaComisionVendedor

        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr
        Dim lStrSql As String
        'Declarar vectores para almacenar los parámetros que se le enviarán al reporte
        Dim aParam(5) As Object
        Dim aValues(5) As Object

        If Not ValidaDatos() Then
            Exit Sub
        End If

        lStrSql = DevuelveQuery()
        gStrSql = lStrSql
        ModEstandar.BorraCmd()
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen datos para el rango de fechas indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            rptVentasSalidaDeMercanciaComisionVendedor.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "Mensaje"
        'aValues(1) = Trim(Me.txtMensaje.Text)
        'aParam(2) = "dDesde"
        'aValues(2) = Me.dtpDesde.Value
        'aParam(3) = "dHasta"
        'aValues(3) = Me.dtpHasta.Value
        'aParam(4) = "Empresa"
        'aValues(4) = Trim(gstrNombCortoEmpresa)

        If (txtMensaje.Text <> Nothing) Then
            pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaComisionVendedor.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaComisionVendedor.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        End If

        If (dtpDesde.Value <> Nothing) Then
            pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaComisionVendedor.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
        End If

        If (dtpHasta.Value <> Nothing) Then
            pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaComisionVendedor.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
        End If

        If (gstrNombCortoEmpresa <> Nothing) Then
            pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
            rptVentasSalidaDeMercanciaComisionVendedor.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
        End If


        frmReportes.reporteActual = rptVentasSalidaDeMercanciaComisionVendedor 'Es el nombre del archivo que se incluyó en el proyecto
        frmReportes.Show()
        ' frmReportes.Imprime(Trim(Me.Text), aParam, aValues)
        Cmd.CommandTimeout = 90

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function ValidaDatos() As Boolean
        If mblnTecleoFechaI Then
            Do While (msglTiempoCambioI) <= 2.1
            Loop
            mblnTecleoFechaI = False
        End If
        If mblnTecleoFechaF Then
            Do While (msglTiempoCambioF) <= 2.1
            Loop
            mblnTecleoFechaF = False
        End If
        System.Windows.Forms.Application.DoEvents()
        Select Case True
            Case Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Unchecked And mintCodSucursal = 0
                MsgBox("Si no quiere imprimir los resultados de todas las sucursales, seleccione una de ellas", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dbcSucursal.Focus()
            Case Me.chkTodosVendedores.CheckState = System.Windows.Forms.CheckState.Unchecked And mintCodVendedor = 0
                MsgBox("Si no quiere imprimir los resultados de todos los vendedores, seleccione uno de ellos", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                Me.dbcVendedor.Focus()
            Case Me.dtpDesde.Value > Me.dtpHasta.Value
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dtpDesde.Focus()
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Private Sub chkTodas_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodas.CheckStateChanged
        Select Case Me.chkTodas.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                Me.dbcSucursal.Text = C_TODAS
                Me.dbcSucursal.Tag = ""
                mintCodSucursal = 0
                Me.dbcSucursal.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                Me.dbcSucursal.Text = ""
                Me.dbcSucursal.Tag = ""
                mintCodSucursal = 0
                Me.dbcSucursal.Enabled = True
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub chkTodosVendedores_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodosVendedores.CheckStateChanged
        Select Case Me.chkTodosVendedores.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                Me.dbcVendedor.Text = C_TODOS
                Me.dbcVendedor.Tag = ""
                mintCodVendedor = 0
                Me.dbcVendedor.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                Me.dbcVendedor.Text = ""
                Me.dbcVendedor.Tag = ""
                mintCodVendedor = 0
                Me.dbcVendedor.Enabled = True
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)
        If Trim(Me.dbcSucursal.Text) = "" Then
            mintCodSucursal = 0
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen WHERE TipoAlmacen = 'P'"
        ModDCombo.DCGotFocus(gStrSql, dbcSucursal)
    End Sub

    Private Sub dbcSucursal_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkTodas.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursal_KeyUpE(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcSucursal.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcSucursal.text)
        '''    If Me.dbcSucursal.SelectedItem <> 0 Then
        '''        dbcSucursal_LostFocus
        '''    End If
        '''    Me.dbcSucursal.text = Aux
    End Sub

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        Dim I As Integer
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        Else
            If Trim(Me.dbcSucursal.Text) = "" Or Trim(Me.dbcSucursal.Text) = C_TODAS Then Exit Sub
        End If
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        Aux = mintCodSucursal
        mintCodSucursal = 0
        ModDCombo.DCLostFocus((Me.dbcSucursal), gStrSql, mintCodSucursal)
    End Sub

    Private Sub dbcSucursal_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcSucursal.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcSucursal.text)
        '''    If Me.dbcSucursal.SelectedItem <> 0 Then
        '''        dbcSucursal_LostFocus
        '''    End If
        '''    Me.dbcSucursal.text = Aux
    End Sub

    Private Sub dbcvendedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcVendedor.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub
        lStrSql = "SELECT codVendedor, LTrim(RTrim(descVendedor)) as descVendedor FROM catVendedores Where descVendedor LIKE '" & Trim(Me.dbcVendedor.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcVendedor))
        If Trim(Me.dbcVendedor.Text) = "" Then
            mintCodVendedor = 0
        End If

Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcvendedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVendedor.Enter
        Pon_Tool()
        gStrSql = "SELECT codVendedor, LTrim(RTrim(descVendedor)) as descVendedor FROM catVendedores order by descVendedor"
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcVendedor))
    End Sub

    Private Sub dbcvendedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcVendedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkTodosVendedores.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcVendedor_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dbcVendedor.KeyUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcVendedor.text)
        '''    If Me.dbcVendedor.SelectedItem <> 0 Then
        '''        dbcvendedor_LostFocus
        '''    End If
        '''    Me.dbcVendedor.text = Aux
    End Sub

    Private Sub dbcvendedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVendedor.Leave
        Dim I As Integer
        Dim Aux As Integer
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        Else
            If Trim(Me.dbcVendedor.Text) = "" Or Trim(Me.dbcVendedor.Text) = C_TODOS Then Exit Sub
        End If
        gStrSql = "SELECT codVendedor, LTrim(RTrim(descVendedor)) as descVendedor FROM catVendedores Where descVendedor LIKE '" & Trim(Me.dbcVendedor.Text) & "%'"
        Aux = mintCodVendedor
        mintCodVendedor = 0
        ModDCombo.DCLostFocus((Me.dbcVendedor), gStrSql, mintCodVendedor)
    End Sub

    Private Sub dbcVendedor_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles dbcVendedor.MouseUp
        '''    Dim Aux As String
        '''    Aux = Trim(Me.dbcVendedor.text)
        '''    If Me.dbcVendedor.SelectedItem <> 0 Then
        '''        dbcvendedor_LostFocus
        '''    End If
        '''    Me.dbcVendedor.text = Aux
    End Sub

    Private Sub dtpDesde_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpDesde.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpDesde_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpDesde.KeyPress
        mblnTecleoFechaI = True
        'msglTiempoCambioI = VB.Timer()
    End Sub

    Private Sub dtpHasta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpHasta.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpHasta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles dtpHasta.KeyPress
        mblnTecleoFechaF = True
        ' msglTiempoCambioF = VB.Timer()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaComisionVendedor_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaComisionVendedor_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaComisionVendedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "CHKTODAS" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaComisionVendedor_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaComisionVendedor_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Me.dtpDesde.MinDate = C_FECHAINICIAL
        Me.dtpDesde.MaxDate = C_FECHAFINAL
        Me.dtpHasta.MinDate = C_FECHAINICIAL
        Me.dtpHasta.MaxDate = C_FECHAFINAL
        Call Me.Nuevo()
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaComisionVendedor_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        If mblnSalir Then
            mblnSalir = False
            Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes 'Sale del Formulario
                    Cancel = 0
                Case MsgBoxResult.No 'No sale del formulario
                    Me.chkTodas.Focus()
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmVtasRPTVentasSalidadeMercanciaComisionVendedor_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        Cmd.CommandTimeout = 90
        frmVtasRPTVentasSalidadeMercanciaPorVendedor = Nothing
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs)

    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkTodosVendedores = New System.Windows.Forms.CheckBox()
        Me.chkTodas = New System.Windows.Forms.CheckBox()
        Me._fraVtas_3 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me._lblVentas_3 = New System.Windows.Forms.Label()
        Me.dbcVendedor = New System.Windows.Forms.ComboBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.fraVtas = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me._fraVtas_3.SuspendLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(6, 205)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(341, 80)
        Me.txtMensaje.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkTodosVendedores
        '
        Me.chkTodosVendedores.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodosVendedores.Checked = True
        Me.chkTodosVendedores.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodosVendedores.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodosVendedores.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodosVendedores.Location = New System.Drawing.Point(6, 54)
        Me.chkTodosVendedores.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodosVendedores.Name = "chkTodosVendedores"
        Me.chkTodosVendedores.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodosVendedores.Size = New System.Drawing.Size(153, 19)
        Me.chkTodosVendedores.TabIndex = 3
        Me.chkTodosVendedores.Text = "Todos los Vendedores"
        Me.chkTodosVendedores.UseVisualStyleBackColor = False
        '
        'chkTodas
        '
        Me.chkTodas.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodas.Checked = True
        Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodas.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodas.Location = New System.Drawing.Point(6, 6)
        Me.chkTodas.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodas.Name = "chkTodas"
        Me.chkTodas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodas.Size = New System.Drawing.Size(160, 22)
        Me.chkTodas.TabIndex = 0
        Me.chkTodas.Text = "Todas las sucursales"
        Me.chkTodas.UseVisualStyleBackColor = False
        '
        '_fraVtas_3
        '
        Me._fraVtas_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_3.Controls.Add(Me.dtpDesde)
        Me._fraVtas_3.Controls.Add(Me.dtpHasta)
        Me._fraVtas_3.Controls.Add(Me._lblVentas_2)
        Me._fraVtas_3.Controls.Add(Me._lblVentas_3)
        Me._fraVtas_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_3.Location = New System.Drawing.Point(11, 115)
        Me._fraVtas_3.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_3.Name = "_fraVtas_3"
        Me._fraVtas_3.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_3.Size = New System.Drawing.Size(336, 54)
        Me._fraVtas_3.TabIndex = 6
        Me._fraVtas_3.TabStop = False
        Me._fraVtas_3.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(74, 20)
        Me.dtpDesde.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(95, 20)
        Me.dtpDesde.TabIndex = 8
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(227, 20)
        Me.dtpHasta.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(95, 20)
        Me.dtpHasta.TabIndex = 10
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_2.Location = New System.Drawing.Point(185, 24)
        Me._lblVentas_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_2.TabIndex = 9
        Me._lblVentas_2.Text = "Hasta"
        '
        '_lblVentas_3
        '
        Me._lblVentas_3.AutoSize = True
        Me._lblVentas_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_3.Location = New System.Drawing.Point(33, 24)
        Me._lblVentas_3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_3.Name = "_lblVentas_3"
        Me._lblVentas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_3.Size = New System.Drawing.Size(38, 13)
        Me._lblVentas_3.TabIndex = 7
        Me._lblVentas_3.Text = "Desde"
        '
        'dbcVendedor
        '
        Me.dbcVendedor.Location = New System.Drawing.Point(64, 79)
        Me.dbcVendedor.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcVendedor.Name = "dbcVendedor"
        Me.dbcVendedor.Size = New System.Drawing.Size(283, 21)
        Me.dbcVendedor.TabIndex = 5
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(66, 29)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(281, 21)
        Me.dbcSucursal.TabIndex = 2
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_1.Location = New System.Drawing.Point(9, 81)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(53, 13)
        Me._lblVentas_1.TabIndex = 4
        Me._lblVentas_1.Text = "Vendedor"
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_0.Location = New System.Drawing.Point(8, 29)
        Me._lblVentas_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(48, 13)
        Me._lblVentas_0.TabIndex = 1
        Me._lblVentas_0.Text = "Sucursal"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblRpt_2.Location = New System.Drawing.Point(9, 180)
        Me._lblRpt_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 11
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(123, 298)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 76
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(8, 298)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 75
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasRPTVentasSalidadeMercanciaComisionVendedor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(359, 341)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.chkTodosVendedores)
        Me.Controls.Add(Me.chkTodas)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._fraVtas_3)
        Me.Controls.Add(Me.dbcVendedor)
        Me.Controls.Add(Me.dbcSucursal)
        Me.Controls.Add(Me._lblVentas_1)
        Me.Controls.Add(Me._lblVentas_0)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 29)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasRPTVentasSalidadeMercanciaComisionVendedor"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Comisiones por Vendedor"
        Me._fraVtas_3.ResumeLayout(False)
        Me._fraVtas_3.PerformLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

End Class