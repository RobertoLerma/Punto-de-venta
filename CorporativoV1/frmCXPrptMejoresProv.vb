Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmCXPrptMejoresProv
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkTodos As System.Windows.Forms.CheckBox
    Public WithEvents dbcProveedor As System.Windows.Forms.ComboBox
    Public WithEvents _fraRpt_4 As System.Windows.Forms.GroupBox
    Public WithEvents chkAcred As System.Windows.Forms.CheckBox
    Public WithEvents chkProv As System.Windows.Forms.CheckBox
    Public WithEvents _fraRpt_3 As System.Windows.Forms.GroupBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents _optMoneda_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optMoneda_2 As System.Windows.Forms.RadioButton
    Public WithEvents _fraRpt_2 As System.Windows.Forms.GroupBox
    Public WithEvents _chkMoneda_2 As System.Windows.Forms.CheckBox
    Public WithEvents _chkMoneda_1 As System.Windows.Forms.CheckBox
    Public WithEvents _chkMoneda_0 As System.Windows.Forms.CheckBox
    Public WithEvents _fraRpt_1 As System.Windows.Forms.GroupBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblRpt_1 As System.Windows.Forms.Label
    Public WithEvents _lblRpt_0 As System.Windows.Forms.Label
    Public WithEvents _fraRpt_0 As System.Windows.Forms.GroupBox
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents chkMoneda As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
    Public WithEvents fraRpt As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents optMoneda As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray

    Const C_DOLARES As Integer = 0
    Const C_PESOS As Integer = 1
    Const C_EUROS As Integer = 2

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean

    Dim cMonedaDeCantidades As String 'Moneda en la que estarán expresadas las cantidades en el reporte
    Dim Tecla As Integer
    Dim mblnFueraChange As Boolean
    Dim mintCodProveedor As Integer
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Dim mblnSalir As Boolean

    Public Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.dtpDesde.Focus()
    End Sub

    Public Sub Nuevo()
        Me.dtpDesde.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = VB6.Format(Today, "dd/MMM/yyyy")
        Me.chkMoneda(C_DOLARES).CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMoneda(C_PESOS).CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMoneda(C_EUROS).CheckState = System.Windows.Forms.CheckState.Checked
        Me.optMoneda(C_DOLARES).Checked = True
        Me.optMoneda(C_PESOS).Checked = False
        Me.optMoneda(C_EUROS).Checked = False
        Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
        Me.txtMensaje.Text = ""
        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
    End Sub

    Public Function DevuelveQuery() As String
        On Error Resume Next
        Dim I As Integer
        Dim cMoneda As String

        Dim cSELECTPROV As String
        Dim cSELECTACRED As String
        Dim cFROMPROV As String
        Dim cFROMACRED As String
        Dim cWHEREPROV As String
        Dim cWHEREACRED As String
        Dim cGROUPBYPROV As String
        Dim cGROUPBYACRED As String
        Dim cORDERBY As String

        'Convertir los totales a la moneda indicada
        If Me.optMoneda(C_DOLARES).Checked Then
            cMoneda = C_DOLAR
            cMonedaDeCantidades = "** Los importes están expresados en Dólares (USD)"
        ElseIf Me.optMoneda(C_PESOS).Checked Then
            cMoneda = C_PESO
            cMonedaDeCantidades = "** Los importes están expresados en Pesos"
        Else
            cMoneda = C_EURO
            cMonedaDeCantidades = "** Los importes están expresados en Euros"
        End If

        'Obtener el query para el proveedor
        If chkProv.CheckState = System.Windows.Forms.CheckState.Checked Then
            'OBTENER EL SELECT
            cSELECTPROV = "select a.codProvAcreed,LTrim(RTrim( b.DescProvAcreed )) as NomProv,Max(FechaCompraEI) as FechaUltimaCompra,Count(FolioOrdenCompra) as NoCompras," & "sum(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.Total, a.TipoCambioC, a.TipoCambioEuroC )) as Total "
            'OBTENER EL FROM
            cFROMPROV = "from OrdenesCompra a, CatProvAcreed b "
            'OBTENER EL WHERE
            cWHEREPROV = "WHERE a.codProvAcreed = b.codProvAcreed and ( a.Estatus = '" & C_STGENERADA & "' or a.Estatus = '" & C_STREGISTRADA & "' )  and (a.FechaCompraEI >= '" & VB6.Format(Me.dtpDesde.Value, "mm/dd/yyyy") & "' and a.FechaCompraEI <= '" & VB6.Format(Me.dtpHasta.Value, "mm/dd/yyyy") & "')"
            If Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                cWHEREPROV = cWHEREPROV & " and a.Moneda <> 'E' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked Then
                cWHEREPROV = cWHEREPROV & " and a.Moneda <> 'P' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                cWHEREPROV = cWHEREPROV & " and a.Moneda = 'D' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                cWHEREPROV = cWHEREPROV & " and a.Moneda = 'P' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked Then
                cWHEREPROV = cWHEREPROV & " and a.Moneda <> 'D' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked Then
                cWHEREPROV = cWHEREPROV & " and a.Moneda = 'E' "
            End If
            If mintCodProveedor <> 0 Then cWHEREPROV = cWHEREPROV & " AND a.codProvAcreed = " & mintCodProveedor
            'OBTENER EL GROUP BY
            cGROUPBYPROV = " group by a.codProvAcreed, b.DescProvAcreed "
        Else
            cSELECTPROV = ""
            cFROMPROV = ""
            cWHEREPROV = ""
            cGROUPBYPROV = ""
        End If

        'Obtener el query para el acreedor
        If chkAcred.CheckState = System.Windows.Forms.CheckState.Checked Then
            'OBTENER EL SELECT
            cSELECTACRED = "select a.codProvAcreed,LTrim(RTrim( b.DescProvAcreed )) as NomProv,Max(FechaRegistro) as FechaUltimaCompra,Count(FolioFactura) as NoCompras," & "sum(dbo.ConvertirCantidad( a.Moneda, '" & cMoneda & "', a.Total, a.TipoCambio, a.TipoCambioEuro )) as Total "
            'OBTENER EL FROM
            cFROMACRED = "from CXPFacturas a, CatProvAcreed b "
            'OBTENER EL WHERE
            cWHEREACRED = "WHERE a.codProvAcreed = b.codProvAcreed and a.Estatus <> 'C' and (a.FechaRegistro >= '" & VB6.Format(Me.dtpDesde.Value, "mm/dd/yyyy") & "' and a.FechaRegistro <= '" & VB6.Format(Me.dtpHasta.Value, "mm/dd/yyyy") & "') and a.TipoFacturaCXP = 'A'"
            If Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                cWHEREACRED = cWHEREACRED & " and a.Moneda <> 'E' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked Then
                cWHEREACRED = cWHEREACRED & " and a.Moneda <> 'P' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                cWHEREACRED = cWHEREACRED & " and a.Moneda = 'D' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked Then
                cWHEREACRED = cWHEREACRED & " and a.Moneda = 'P' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Checked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked Then
                cWHEREACRED = cWHEREACRED & " and a.Moneda <> 'D' "
            ElseIf Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Checked Then
                cWHEREACRED = cWHEREACRED & " and a.Moneda = 'E' "
            End If
            If mintCodProveedor <> 0 Then cWHEREACRED = cWHEREACRED & " AND a.codProvAcreed = " & mintCodProveedor
            'OBTENER EL GROUP BY
            cGROUPBYACRED = " group by a.codProvAcreed, b.DescProvAcreed "
        Else
            cSELECTACRED = ""
            cFROMACRED = ""
            cWHEREACRED = ""
            cGROUPBYACRED = ""
        End If

        'Obtenemos el Query Final
        DevuelveQuery = "SELECT * FROM (" & cSELECTPROV & cFROMPROV & cWHEREPROV & cGROUPBYPROV & IIf(chkProv.CheckState = System.Windows.Forms.CheckState.Checked And chkAcred.CheckState = System.Windows.Forms.CheckState.Checked, " UNION ", "") & cSELECTACRED & cFROMACRED & cWHEREACRED & cGROUPBYACRED & ") RES ORDER BY Total Desc"
    End Function

    Public Sub Imprime()

        Dim rptCXPrptMejoresProv As New rptCXPrptMejoresProv
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        'On Error GoTo MErr

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
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute
        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existen datos para el rango de fechas indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        Else
            rptCXPrptMejoresProv.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "Mensaje"
        'aValues(1) = Trim(Me.txtMensaje.Text)
        'aParam(2) = "dDesde"
        'aValues(2) = Me.dtpDesde.Value
        'aParam(3) = "dHasta"
        'aValues(3) = Me.dtpHasta.Value
        'aParam(4) = "MonedaDeCantidades"
        'aValues(4) = Trim(cMonedaDeCantidades)
        'aParam(5) = "Empresa"
        'aValues(5) = Trim(gstrNombCortoEmpresa)

        If (txtMensaje.Text <> Nothing Or txtMensaje.Text <> "") Then
            pdvNum.Value = txtMensaje.Text : pvNum.Add(pdvNum)
            rptCXPrptMejoresProv.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        Else
            pdvNum.Value = "" : pvNum.Add(pdvNum)
            rptCXPrptMejoresProv.DataDefinition.ParameterFields("Mensaje").ApplyCurrentValues(pvNum)
        End If

        If (dtpDesde.Value <> Nothing) Then
            pdvNum.Value = dtpDesde.Value : pvNum.Add(pdvNum)
            rptCXPrptMejoresProv.DataDefinition.ParameterFields("dDesde").ApplyCurrentValues(pvNum)
        End If

        If (dtpHasta.Value <> Nothing) Then
            pdvNum.Value = dtpHasta.Value : pvNum.Add(pdvNum)
            rptCXPrptMejoresProv.DataDefinition.ParameterFields("dHasta").ApplyCurrentValues(pvNum)
        End If

        If (cMonedaDeCantidades <> Nothing Or cMonedaDeCantidades <> "") Then
            pdvNum.Value = cMonedaDeCantidades : pvNum.Add(pdvNum)
            rptCXPrptMejoresProv.DataDefinition.ParameterFields("MonedaDeCantidades").ApplyCurrentValues(pvNum)
        End If

        If (gstrNombCortoEmpresa <> Nothing Or gstrNombCortoEmpresa <> "") Then
            pdvNum.Value = gstrNombCortoEmpresa : pvNum.Add(pdvNum)
            rptCXPrptMejoresProv.DataDefinition.ParameterFields("Empresa").ApplyCurrentValues(pvNum)
        End If


        'frmReportes.Report = rptCXPrptMejoresProv 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime(Trim(Me.Text), aParam, aValues)
        frmReportes.reporteActual = rptCXPrptMejoresProv
        frmReportes.Show()

MErr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Public Function ValidaDatos() As Boolean
        If mblnTecleoFechaI Then
            Do While (VB.Timer() - msglTiempoCambioI) <= 2.1
            Loop
            mblnTecleoFechaI = False
        End If
        If mblnTecleoFechaF Then
            Do While (VB.Timer() - msglTiempoCambioF) <= 2.1
            Loop
            mblnTecleoFechaF = False
        End If
        System.Windows.Forms.Application.DoEvents()
        Select Case True
            Case Me.dtpDesde.Value > Me.dtpHasta.Value
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dtpDesde.Focus()
                Exit Function
            Case Me.chkMoneda(0).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(1).CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkMoneda(2).CheckState = System.Windows.Forms.CheckState.Unchecked
                MsgBox("Debe seleccionar por lo menos un tipo de moneda", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.chkMoneda(0).Focus()
                Exit Function
            Case chkTodos.CheckState = System.Windows.Forms.CheckState.Unchecked And mintCodProveedor = 0
                MsgBox("Debe Seleccionar un Proveedor/Acreedor", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                dbcProveedor.Focus()
                Exit Function
            Case Else
                ValidaDatos = True
        End Select
    End Function

    Private Sub chkAcred_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAcred.CheckStateChanged
        Select Case True
            Case Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
                'Me.chkTodos.Enabled = False

            Case Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.Enabled = True

            Case Me.chkProv.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.Enabled = True

            Case Me.chkProv.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me.chkTodos.Enabled = False

        End Select
        mblnFueraChange = True
        mintCodProveedor = 0
        Me.dbcProveedor.Text = "[ Todos ... ]"
        Me.dbcProveedor.Tag = ""
        Me.dbcProveedor.Enabled = False
        mblnFueraChange = False
    End Sub

    Private Sub chkAcred_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAcred.Enter
        Pon_Tool()
    End Sub

    Private Sub chkMoneda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkMoneda.Enter
        Dim Index As Integer = chkMoneda.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub chkProv_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkProv.CheckStateChanged
        Select Case True
            Case Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
                'Me.chkTodos.Enabled = False

            Case Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.Enabled = True

            Case Me.chkProv.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked
                Me.chkTodos.Enabled = True

            Case Me.chkProv.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Unchecked
                Me.chkTodos.Enabled = False

        End Select
        mblnFueraChange = True
        mintCodProveedor = 0
        Me.dbcProveedor.Text = "[ Todos ... ]"
        Me.dbcProveedor.Tag = ""
        Me.dbcProveedor.Enabled = False
        mblnFueraChange = False
    End Sub

    Private Sub chkProv_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkProv.Enter
        Pon_Tool()
    End Sub

    Private Sub chkTodos_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodos.CheckStateChanged
        If Me.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
            mblnFueraChange = True
            mintCodProveedor = 0
            Me.dbcProveedor.Text = "[ Todos ... ]"
            Me.dbcProveedor.Tag = ""
            mblnFueraChange = False
            Me.dbcProveedor.Enabled = False
        Else
            mblnFueraChange = True
            mintCodProveedor = 0
            Me.dbcProveedor.Text = ""
            Me.dbcProveedor.Tag = ""
            mblnFueraChange = False
            Me.dbcProveedor.Enabled = True
        End If
    End Sub

    Private Sub chkTodos_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodos.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcProveedor_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.CursorChanged
        On Error GoTo MErr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        If Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And chkAcred.CheckState = System.Windows.Forms.CheckState.Checked Then
            lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ElseIf Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And chkAcred.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ElseIf Me.chkProv.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked Then
            lStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TACREEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        End If
        ModDCombo.DCChange(lStrSql, Tecla, dbcProveedor)
        If Trim(Me.dbcProveedor.Text) = "" Then
            dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        End If
MErr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcProveedor_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Enter
        Pon_Tool()
        If Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And chkAcred.CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed ORDER BY descProvAcreed"
        ElseIf Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And chkAcred.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & C_TPROVEEDOR & "' ORDER BY descProvAcreed"
        ElseIf Me.chkProv.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed WHERE Tipo = '" & C_TACREEDOR & "' ORDER BY descProvAcreed"
        End If
        ModDCombo.DCGotFocus(gStrSql, dbcProveedor)
    End Sub

    Private Sub dbcProveedor_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcProveedor.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkTodos.Focus()
        End If
        Tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcProveedor_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcProveedor.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        If Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ElseIf Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TPROVEEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        ElseIf Me.chkProv.CheckState = System.Windows.Forms.CheckState.Unchecked And Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT codProvAcreed, LTrim(RTrim(descProvAcreed)) as descProvAcreed FROM catProvAcreed Where Tipo = '" & C_TACREEDOR & "' and descProvAcreed LIKE '" & Trim(Me.dbcProveedor.Text) & "%'"
        End If
        Aux = mintCodProveedor
        mintCodProveedor = 0
        ModDCombo.DCLostFocus(dbcProveedor, gStrSql, mintCodProveedor)
    End Sub

    Private Sub dbcProveedor_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcProveedor.MouseUp
        Dim Aux As String
        Aux = Trim(Me.dbcProveedor.Text)
        'If Me.dbcProveedor.SelectedItem <> 0 Then
        '    dbcProveedor_Leave(dbcProveedor, New System.EventArgs())
        'End If
        Me.dbcProveedor.Text = Aux
    End Sub

    Private Sub dtpDesde_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpDesde.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpDesde_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpDesde.KeyPress
        mblnTecleoFechaI = True
        msglTiempoCambioI = VB.Timer()
    End Sub

    Private Sub dtpHasta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpHasta.Enter
        Pon_Tool()
    End Sub

    Private Sub dtpHasta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As KeyPressEventArgs) Handles dtpHasta.KeyPress
        mblnTecleoFechaF = True
        msglTiempoCambioF = VB.Timer()
    End Sub

    Private Sub frmCXPrptMejoresProv_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmCXPrptMejoresProv_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmCXPrptMejoresProv_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                If UCase(Me.ActiveControl.Name) = "CHKPROV" Then
                    mblnSalir = True
                    Me.Close()
                Else
                    ModEstandar.RetrocederTab(Me)
                End If
        End Select
    End Sub

    Private Sub frmCXPrptMejoresProv_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmCXPrptMejoresProv_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmCXPrptMejoresProv_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim Cancel As Boolean = eventArgs.Cancel
        'Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'If mblnSalir Then
        '    mblnSalir = False
        '    Select Case MsgBox("¿Desea abandonar el proceso?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
        '        Case MsgBoxResult.Yes 'Sale del Formulario
        '            Cancel = 0
        '        Case MsgBoxResult.No 'No sale del formulario
        '            Me.chkProv.Focus()
        '            Cancel = 1
        '    End Select
        'End If
        'eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmCXPrptMejoresProv_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub optMoneda_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optMoneda.Enter
        Dim Index As Integer = optMoneda.GetIndex(eventSender)
        Pon_Tool()
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkTodos = New System.Windows.Forms.CheckBox()
        Me.chkAcred = New System.Windows.Forms.CheckBox()
        Me.chkProv = New System.Windows.Forms.CheckBox()
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me._optMoneda_0 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_1 = New System.Windows.Forms.RadioButton()
        Me._optMoneda_2 = New System.Windows.Forms.RadioButton()
        Me._chkMoneda_2 = New System.Windows.Forms.CheckBox()
        Me._chkMoneda_1 = New System.Windows.Forms.CheckBox()
        Me._chkMoneda_0 = New System.Windows.Forms.CheckBox()
        Me._fraRpt_4 = New System.Windows.Forms.GroupBox()
        Me.dbcProveedor = New System.Windows.Forms.ComboBox()
        Me._fraRpt_3 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_2 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_1 = New System.Windows.Forms.GroupBox()
        Me._fraRpt_0 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblRpt_1 = New System.Windows.Forms.Label()
        Me._lblRpt_0 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.chkMoneda = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(Me.components)
        Me.fraRpt = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.optMoneda = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me._fraRpt_4.SuspendLayout()
        Me._fraRpt_3.SuspendLayout()
        Me._fraRpt_2.SuspendLayout()
        Me._fraRpt_1.SuspendLayout()
        Me._fraRpt_0.SuspendLayout()
        CType(Me.chkMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkTodos
        '
        Me.chkTodos.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTodos.Location = New System.Drawing.Point(16, 28)
        Me.chkTodos.Name = "chkTodos"
        Me.chkTodos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodos.Size = New System.Drawing.Size(58, 17)
        Me.chkTodos.TabIndex = 2
        Me.chkTodos.Text = "Todos"
        Me.ToolTip1.SetToolTip(Me.chkTodos, "Selecciona todos los proveedores")
        Me.chkTodos.UseVisualStyleBackColor = False
        '
        'chkAcred
        '
        Me.chkAcred.BackColor = System.Drawing.SystemColors.Control
        Me.chkAcred.Checked = True
        Me.chkAcred.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAcred.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAcred.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkAcred.Location = New System.Drawing.Point(200, 24)
        Me.chkAcred.Name = "chkAcred"
        Me.chkAcred.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAcred.Size = New System.Drawing.Size(137, 27)
        Me.chkAcred.TabIndex = 1
        Me.chkAcred.Text = "Acreedores"
        Me.ToolTip1.SetToolTip(Me.chkAcred, "Facturas de Gastos")
        Me.chkAcred.UseVisualStyleBackColor = False
        '
        'chkProv
        '
        Me.chkProv.BackColor = System.Drawing.SystemColors.Control
        Me.chkProv.Checked = True
        Me.chkProv.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkProv.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkProv.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkProv.Location = New System.Drawing.Point(42, 24)
        Me.chkProv.Name = "chkProv"
        Me.chkProv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkProv.Size = New System.Drawing.Size(153, 27)
        Me.chkProv.TabIndex = 0
        Me.chkProv.Text = "Proveedores"
        Me.ToolTip1.SetToolTip(Me.chkProv, "Facturas de Compras")
        Me.chkProv.UseVisualStyleBackColor = False
        '
        'txtMensaje
        '
        Me.txtMensaje.AcceptsReturn = True
        Me.txtMensaje.BackColor = System.Drawing.SystemColors.Window
        Me.txtMensaje.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMensaje.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMensaje.Location = New System.Drawing.Point(11, 311)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(345, 71)
        Me.txtMensaje.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        '_optMoneda_0
        '
        Me._optMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_0.Checked = True
        Me._optMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_0, CType(0, Short))
        Me._optMoneda_0.Location = New System.Drawing.Point(24, 24)
        Me._optMoneda_0.Name = "_optMoneda_0"
        Me._optMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_0.Size = New System.Drawing.Size(70, 17)
        Me._optMoneda_0.TabIndex = 9
        Me._optMoneda_0.TabStop = True
        Me._optMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._optMoneda_0, "Los importes del reporte aparecerán en dólares")
        Me._optMoneda_0.UseVisualStyleBackColor = False
        '
        '_optMoneda_1
        '
        Me._optMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_1, CType(1, Short))
        Me._optMoneda_1.Location = New System.Drawing.Point(24, 48)
        Me._optMoneda_1.Name = "_optMoneda_1"
        Me._optMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_1.Size = New System.Drawing.Size(70, 17)
        Me._optMoneda_1.TabIndex = 10
        Me._optMoneda_1.TabStop = True
        Me._optMoneda_1.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._optMoneda_1, "Los importes del reporte aparecerán en Pesos")
        Me._optMoneda_1.UseVisualStyleBackColor = False
        '
        '_optMoneda_2
        '
        Me._optMoneda_2.BackColor = System.Drawing.SystemColors.Control
        Me._optMoneda_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optMoneda.SetIndex(Me._optMoneda_2, CType(2, Short))
        Me._optMoneda_2.Location = New System.Drawing.Point(24, 72)
        Me._optMoneda_2.Name = "_optMoneda_2"
        Me._optMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optMoneda_2.Size = New System.Drawing.Size(70, 19)
        Me._optMoneda_2.TabIndex = 11
        Me._optMoneda_2.TabStop = True
        Me._optMoneda_2.Text = "Euros"
        Me.ToolTip1.SetToolTip(Me._optMoneda_2, "Los importes del reporte aparecerán en Euros")
        Me._optMoneda_2.UseVisualStyleBackColor = False
        '
        '_chkMoneda_2
        '
        Me._chkMoneda_2.BackColor = System.Drawing.SystemColors.Control
        Me._chkMoneda_2.Checked = True
        Me._chkMoneda_2.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkMoneda_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMoneda.SetIndex(Me._chkMoneda_2, CType(2, Short))
        Me._chkMoneda_2.Location = New System.Drawing.Point(24, 72)
        Me._chkMoneda_2.Name = "_chkMoneda_2"
        Me._chkMoneda_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_2.Size = New System.Drawing.Size(81, 17)
        Me._chkMoneda_2.TabIndex = 8
        Me._chkMoneda_2.Text = "Euros"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_2, "Selecciona todas las compras en Euros")
        Me._chkMoneda_2.UseVisualStyleBackColor = False
        '
        '_chkMoneda_1
        '
        Me._chkMoneda_1.BackColor = System.Drawing.SystemColors.Control
        Me._chkMoneda_1.Checked = True
        Me._chkMoneda_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkMoneda_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMoneda.SetIndex(Me._chkMoneda_1, CType(1, Short))
        Me._chkMoneda_1.Location = New System.Drawing.Point(24, 48)
        Me._chkMoneda_1.Name = "_chkMoneda_1"
        Me._chkMoneda_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_1.Size = New System.Drawing.Size(81, 17)
        Me._chkMoneda_1.TabIndex = 7
        Me._chkMoneda_1.Text = "Pesos"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_1, "Selecciona todas las compras en Pesos")
        Me._chkMoneda_1.UseVisualStyleBackColor = False
        '
        '_chkMoneda_0
        '
        Me._chkMoneda_0.BackColor = System.Drawing.SystemColors.Control
        Me._chkMoneda_0.Checked = True
        Me._chkMoneda_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._chkMoneda_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._chkMoneda_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMoneda.SetIndex(Me._chkMoneda_0, CType(0, Short))
        Me._chkMoneda_0.Location = New System.Drawing.Point(24, 24)
        Me._chkMoneda_0.Name = "_chkMoneda_0"
        Me._chkMoneda_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._chkMoneda_0.Size = New System.Drawing.Size(81, 17)
        Me._chkMoneda_0.TabIndex = 6
        Me._chkMoneda_0.Text = "Dólares"
        Me.ToolTip1.SetToolTip(Me._chkMoneda_0, "Selecciona todas las compras en Dólares")
        Me._chkMoneda_0.UseVisualStyleBackColor = False
        '
        '_fraRpt_4
        '
        Me._fraRpt_4.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_4.Controls.Add(Me.chkTodos)
        Me._fraRpt_4.Controls.Add(Me.dbcProveedor)
        Me._fraRpt_4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_4, CType(4, Short))
        Me._fraRpt_4.Location = New System.Drawing.Point(11, 63)
        Me._fraRpt_4.Name = "_fraRpt_4"
        Me._fraRpt_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_4.Size = New System.Drawing.Size(345, 65)
        Me._fraRpt_4.TabIndex = 20
        Me._fraRpt_4.TabStop = False
        Me._fraRpt_4.Text = "Proveedor"
        '
        'dbcProveedor
        '
        Me.dbcProveedor.Location = New System.Drawing.Point(80, 24)
        Me.dbcProveedor.Name = "dbcProveedor"
        Me.dbcProveedor.Size = New System.Drawing.Size(257, 21)
        Me.dbcProveedor.TabIndex = 3
        '
        '_fraRpt_3
        '
        Me._fraRpt_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_3.Controls.Add(Me.chkAcred)
        Me._fraRpt_3.Controls.Add(Me.chkProv)
        Me._fraRpt_3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_3, CType(3, Short))
        Me._fraRpt_3.Location = New System.Drawing.Point(11, 6)
        Me._fraRpt_3.Name = "_fraRpt_3"
        Me._fraRpt_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_3.Size = New System.Drawing.Size(345, 57)
        Me._fraRpt_3.TabIndex = 19
        Me._fraRpt_3.TabStop = False
        '
        '_fraRpt_2
        '
        Me._fraRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_2.Controls.Add(Me._optMoneda_0)
        Me._fraRpt_2.Controls.Add(Me._optMoneda_1)
        Me._fraRpt_2.Controls.Add(Me._optMoneda_2)
        Me._fraRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_2, CType(2, Short))
        Me._fraRpt_2.Location = New System.Drawing.Point(187, 192)
        Me._fraRpt_2.Name = "_fraRpt_2"
        Me._fraRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_2.Size = New System.Drawing.Size(169, 97)
        Me._fraRpt_2.TabIndex = 17
        Me._fraRpt_2.TabStop = False
        Me._fraRpt_2.Text = "Presentar en ..."
        '
        '_fraRpt_1
        '
        Me._fraRpt_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_1.Controls.Add(Me._chkMoneda_2)
        Me._fraRpt_1.Controls.Add(Me._chkMoneda_1)
        Me._fraRpt_1.Controls.Add(Me._chkMoneda_0)
        Me._fraRpt_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_1, CType(1, Short))
        Me._fraRpt_1.Location = New System.Drawing.Point(8, 192)
        Me._fraRpt_1.Name = "_fraRpt_1"
        Me._fraRpt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_1.Size = New System.Drawing.Size(169, 97)
        Me._fraRpt_1.TabIndex = 16
        Me._fraRpt_1.TabStop = False
        Me._fraRpt_1.Text = "Moneda"
        '
        '_fraRpt_0
        '
        Me._fraRpt_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraRpt_0.Controls.Add(Me.dtpDesde)
        Me._fraRpt_0.Controls.Add(Me.dtpHasta)
        Me._fraRpt_0.Controls.Add(Me._lblRpt_1)
        Me._fraRpt_0.Controls.Add(Me._lblRpt_0)
        Me._fraRpt_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRpt.SetIndex(Me._fraRpt_0, CType(0, Short))
        Me._fraRpt_0.Location = New System.Drawing.Point(11, 130)
        Me._fraRpt_0.Name = "_fraRpt_0"
        Me._fraRpt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraRpt_0.Size = New System.Drawing.Size(345, 57)
        Me._fraRpt_0.TabIndex = 13
        Me._fraRpt_0.TabStop = False
        Me._fraRpt_0.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(72, 21)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(97, 20)
        Me.dtpDesde.TabIndex = 4
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(232, 21)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(97, 20)
        Me.dtpHasta.TabIndex = 5
        '
        '_lblRpt_1
        '
        Me._lblRpt_1.AutoSize = True
        Me._lblRpt_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRpt.SetIndex(Me._lblRpt_1, CType(1, Short))
        Me._lblRpt_1.Location = New System.Drawing.Point(184, 25)
        Me._lblRpt_1.Name = "_lblRpt_1"
        Me._lblRpt_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_1.Size = New System.Drawing.Size(44, 13)
        Me._lblRpt_1.TabIndex = 15
        Me._lblRpt_1.Text = "hasta el"
        '
        '_lblRpt_0
        '
        Me._lblRpt_0.AutoSize = True
        Me._lblRpt_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRpt.SetIndex(Me._lblRpt_0, CType(0, Short))
        Me._lblRpt_0.Location = New System.Drawing.Point(16, 25)
        Me._lblRpt_0.Name = "_lblRpt_0"
        Me._lblRpt_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_0.Size = New System.Drawing.Size(49, 13)
        Me._lblRpt_0.TabIndex = 14
        Me._lblRpt_0.Text = "Desde el"
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblRpt.SetIndex(Me._lblRpt_2, CType(2, Short))
        Me._lblRpt_2.Location = New System.Drawing.Point(11, 297)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 18
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'chkMoneda
        '
        '
        'optMoneda
        '
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(130, 401)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 118
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(15, 401)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 117
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(245, 402)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 116
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmCXPrptMejoresProv
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(365, 449)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me._fraRpt_4)
        Me.Controls.Add(Me._fraRpt_3)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me._fraRpt_2)
        Me.Controls.Add(Me._fraRpt_1)
        Me.Controls.Add(Me._fraRpt_0)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(281, 130)
        Me.MaximizeBox = False
        Me.Name = "frmCXPrptMejoresProv"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Reporte de los mejores Proveedores y Acreedores"
        Me._fraRpt_4.ResumeLayout(False)
        Me._fraRpt_3.ResumeLayout(False)
        Me._fraRpt_2.ResumeLayout(False)
        Me._fraRpt_1.ResumeLayout(False)
        Me._fraRpt_0.ResumeLayout(False)
        Me._fraRpt_0.PerformLayout()
        CType(Me.chkMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optMoneda, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub
End Class