Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frmRptExistenciasyCostos
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents Frame5 As System.Windows.Forms.GroupBox
    Public WithEvents optTotal As System.Windows.Forms.RadioButton
    Public WithEvents optxSuc As System.Windows.Forms.RadioButton
    Public WithEvents chkRangoSinExistencia As System.Windows.Forms.CheckBox
    Public WithEvents chkRangoTodos As System.Windows.Forms.CheckBox
    Public WithEvents txtRangoDesde As System.Windows.Forms.TextBox
    Public WithEvents txtRangoHasta As System.Windows.Forms.TextBox
    Public WithEvents lblHasta As System.Windows.Forms.Label
    Public WithEvents lblDesde As System.Windows.Forms.Label
    Public WithEvents fraRangoExistencia As System.Windows.Forms.GroupBox
    Public WithEvents cmbOrdArticulo As System.Windows.Forms.ComboBox
    Public WithEvents cmbOrdExistencia As System.Windows.Forms.ComboBox
    Public WithEvents optOrdExistencia As System.Windows.Forms.RadioButton
    Public WithEvents optOrdArticulo As System.Windows.Forms.RadioButton
    Public WithEvents fraOrdenamiento As System.Windows.Forms.GroupBox
    Public WithEvents chkCostoFactura As System.Windows.Forms.CheckBox
    Public WithEvents chkCostoIndirecto As System.Windows.Forms.CheckBox
    Public WithEvents chkCostoAdicional As System.Windows.Forms.CheckBox
    Public WithEvents chkIncluirIVA As System.Windows.Forms.CheckBox
    Public WithEvents optAlCosto As System.Windows.Forms.RadioButton
    Public WithEvents optPrecioPublico As System.Windows.Forms.RadioButton
    Public WithEvents optUltimoCostoPesos As System.Windows.Forms.RadioButton
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    Public WithEvents chkMostrarAparatdos As System.Windows.Forms.CheckBox
    Public WithEvents chkRelojeria As System.Windows.Forms.CheckBox
    Public WithEvents chkVarios As System.Windows.Forms.CheckBox
    Public WithEvents chkJoyeria As System.Windows.Forms.CheckBox
    Public WithEvents _Frame3_0 As System.Windows.Forms.GroupBox
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents dbcJFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcJLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcJSubLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcVLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcRMarca As System.Windows.Forms.ComboBox
    Public WithEvents dbcRModelo As System.Windows.Forms.ComboBox
    Public WithEvents dbcVFamilia As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_8 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_7 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_6 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_4 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_3 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents fraGrupo As System.Windows.Forms.GroupBox
    Public WithEvents txtCodOrigen As System.Windows.Forms.TextBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents dbcOrigen1 As System.Windows.Forms.ComboBox
    Public WithEvents dtpFechaCorte As System.Windows.Forms.DateTimePicker
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Frame3 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray



    Dim mblnSalir As Boolean
    Dim mintJFamilia As Integer
    Dim mintJLinea As Integer
    Dim mintJSubLinea As Integer
    Dim mintVFamilia As Integer
    Dim mintVLinea As Integer
    Dim mintRMarca As Integer
    Dim mintRModelo As Integer
    Dim intCodSucursal As Integer
    Dim intCodOrigen As Integer
    Dim mblnFueraChange As Boolean
    Dim tecla As Integer
    Dim mstrParamJFamilia As String ' Es el Parámetro que se enviará al Reporte de Crystal para que muestre los datos.
    Dim mstrParamJLinea As String
    Dim mstrParamJSubLinea As String
    Dim mstrParamVFamilia As String ' Es el Parámetro que se enviará al Reporte de Crystal para que muestre los datos.
    Dim mstrParamVLinea As String
    Dim mstrParamMarca As String
    Dim mstrParamModelo As String

    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Friend WithEvents btnBuscar As Button
    Const C_NINGUNA As String = "[ Vacío ... ]"

    Sub Imprime()
        Dim rptExistenciasyCostos_Total As New rptExistenciasyCostos_Total
        Dim rptExistenciasyCostos As New rptExistenciasyCostos
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr

        Dim aParam(3) As Object
        Dim aValues(3) As Object
        Dim FechaInicio As Date
        Dim FechaFin As Date
        Dim TextoAdicional As String
        Dim Encabezado As String
        Dim ConsultaGuardar As String
        Dim ConsultaReporte As String
        Dim mblnTRansaccion As Boolean
        If ValidaDatos() = False Then Exit Sub
        Dim MostrarApartados As Boolean

        Encabezado = "Existencias y Costos en Almacén"
        If chkMostrarAparatdos.CheckState = System.Windows.Forms.CheckState.Checked Then
            MostrarApartados = True
        Else
            MostrarApartados = False
        End If

        gStrSql = DevuelveQuery()
        If Trim(gStrSql) = "" Then Exit Sub

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existe que reportar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            Exit Sub
        Else
            rptExistenciasyCostos.SetDataSource(frmReportes.rsReport)
            rptExistenciasyCostos_Total.SetDataSource(frmReportes.rsReport)
        End If


        'aParam(1) = "EncabezadoReporte"
        'aValues(1) = Encabezado
        'aParam(2) = "FechaCorte"
        'aValues(2) = dtpFechaCorte.Value
        'aParam(3) = "MostrarApartados"
        'aValues(3) = MostrarApartados

        If (Encabezado <> Nothing) Then
            pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
            rptExistenciasyCostos.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
        End If

        If (dtpFechaCorte.Value <> Nothing) Then
            pdvNum.Value = dtpFechaCorte.Value : pvNum.Add(pdvNum)
            rptExistenciasyCostos.DataDefinition.ParameterFields("FechaCorte").ApplyCurrentValues(pvNum)
        End If

        If (MostrarApartados <> Nothing) Then
            pdvNum.Value = MostrarApartados : pvNum.Add(pdvNum)
            rptExistenciasyCostos.DataDefinition.ParameterFields("MostrarApartados").ApplyCurrentValues(pvNum)
        End If


        'Es el nombre del archivo que se incluyó en el proyecto
        If optxSuc.Checked Then
            If optAlCosto.Checked Then
                'rptExistenciasyCostos.TXTIMPORTE.SetText("COSTO") 
                'Dim txt As CrystalDecisions.CrystalReports.Engine.TextObject
                'txt = rptExistenciasyCostos.ReportDefinition.ReportObjects("TXTIMPORTE")
                Dim objText As CrystalDecisions.CrystalReports.Engine.TextObject = rptExistenciasyCostos.ReportDefinition.Sections(1).ReportObjects("TXTIMPORTE")
                objText.Text = "COSTO"
                'txt.Text = textbox1.text
                'pdvNum.Value = TXTIMPORTE : pvNum.Add(pdvNum)
                'rptExistenciasyCostos.DataDefinition.ParameterFields("COSTO").ApplyCurrentValues(pvNum)
            ElseIf optPrecioPublico.Checked Then
                'rptExistenciasyCostos.TXTIMPORTE.SetText("PRECIO PUB")
            ElseIf optUltimoCostoPesos.Checked Then
                'rptExistenciasyCostos.TXTIMPORTE.SetText("COSTO $")
            End If

            'frmReportes.Report = rptExistenciasyCostos
            'rptExistenciasyCostos.SetDataSource(frmReportes.rsReport)
            frmReportes.reporteActual = rptExistenciasyCostos
            frmReportes.Show()

        Else

            If (Encabezado <> Nothing) Then
                pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
                rptExistenciasyCostos_Total.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
            End If

            If (dtpFechaCorte.Value <> Nothing) Then
                pdvNum.Value = dtpFechaCorte.Value : pvNum.Add(pdvNum)
                rptExistenciasyCostos_Total.DataDefinition.ParameterFields("FechaCorte").ApplyCurrentValues(pvNum)
            End If

            If (MostrarApartados <> Nothing) Then
                pdvNum.Value = MostrarApartados : pvNum.Add(pdvNum)
                rptExistenciasyCostos_Total.DataDefinition.ParameterFields("MostrarApartados").ApplyCurrentValues(pvNum)
            End If

            If optAlCosto.Checked Then
                '    rptExistenciasyCostos_Total.TXTIMPORTE.SetText("A ULTIMO COSTO")

                'ElseIf optPrecioPublico.Checked Then
                '    If chkIncluirIVA.CheckState = 1 Then
                '        rptExistenciasyCostos_Total.TXTIMPORTE.SetText("A PRECIO PUBLICO C/IVA")
                '    Else
                '        rptExistenciasyCostos_Total.TXTIMPORTE.SetText("A PRECIO PUBLICO S/IVA")
                '    End If
                'ElseIf optUltimoCostoPesos.Checked Then
                '    rptExistenciasyCostos_Total.TXTIMPORTE.SetText("COSTO EN PESOS")
            End If

            'frmReportes.Report = rptExistenciasyCostos_Total 
            'rptExistenciasyCostos_Total.Text12.SetText("INVENTARIO TOTAL")

            'rptExistenciasyCostos_Total.DataDefinition.FormulaFields("Text12").Text = "INVENTARIO TOTAL"


            frmReportes.reporteActual = rptExistenciasyCostos_Total
            frmReportes.Show()
            'frmReportes.Imprime(Me.Text, aParam, aValues)
        End If
        Exit Sub

Merr:
        If mblnTRansaccion = True Then Cnn.RollbackTrans()
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub



    Function ValidaDatos() As Boolean
        If optxSuc.Checked Then
            If Trim(dbcSucursales.Text) = "" Then
                MsgBox("Debe proporcionar la sucursal para generar el reporte.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                dbcSucursales.Focus()
                Exit Function
            End If
            If chkJoyeria.CheckState = System.Windows.Forms.CheckState.Unchecked And chkRelojeria.CheckState = System.Windows.Forms.CheckState.Unchecked And chkVarios.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                MsgBox("Debe seleccionar un grupo para generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                Exit Function
            End If
        End If

        If optUltimoCostoPesos.Checked = True And chkCostoFactura.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            MsgBox("Si desea generar el reporte con últimos costos, es necesario que seleccione el Costo Factura.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            chkCostoFactura.Focus()
            Exit Function
        End If
        ValidaDatos = True
    End Function

    Private Sub chkJoyeria_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkJoyeria.CheckStateChanged
        Select Case Me.chkJoyeria.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                mintJFamilia = 0
                Me.dbcJFamilia.Text = C_TODAS
                Me.dbcJFamilia.Enabled = True
                mintJLinea = 0
                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = False
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                mintJFamilia = 0
                Me.dbcJFamilia.Text = C_NINGUNA
                Me.dbcJFamilia.Enabled = False
                mintJLinea = 0
                Me.dbcJLinea.Text = C_NINGUNA
                Me.dbcJLinea.Enabled = False
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_NINGUNA
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub chkRangoTodos_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRangoTodos.CheckStateChanged
        If chkRangoTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
            chkRangoSinExistencia.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRangoSinExistencia.Enabled = True
            txtRangoDesde.Enabled = False
            txtRangoHasta.Enabled = False
            lblDesde.Enabled = False
            lblHasta.Enabled = False
        ElseIf chkRangoTodos.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            chkRangoSinExistencia.Enabled = False
            chkRangoSinExistencia.CheckState = System.Windows.Forms.CheckState.Unchecked
            txtRangoDesde.Enabled = True
            txtRangoHasta.Enabled = True
            lblDesde.Enabled = True
            lblHasta.Enabled = True
        End If
    End Sub

    Private Sub chkRelojeria_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRelojeria.CheckStateChanged
        Select Case Me.chkRelojeria.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                mintRMarca = 0
                Me.dbcRMarca.Text = C_TODAS
                Me.dbcRMarca.Enabled = True
                mintRModelo = 0
                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                mintRMarca = 0
                Me.dbcRMarca.Text = C_NINGUNA
                Me.dbcRMarca.Enabled = False
                mintRModelo = 0
                Me.dbcRModelo.Text = C_NINGUNA
                Me.dbcRModelo.Enabled = False
                mblnFueraChange = False
        End Select
    End Sub


    Private Sub chkVarios_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkVarios.CheckStateChanged
        Select Case Me.chkVarios.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                mintVFamilia = 0
                Me.dbcVFamilia.Text = C_TODAS
                Me.dbcVFamilia.Enabled = True
                mintVLinea = 0
                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = False
                mblnFueraChange = False
            Case Else
                mblnFueraChange = True
                mintVFamilia = 0
                Me.dbcVFamilia.Text = C_NINGUNA
                Me.dbcVFamilia.Enabled = False
                mintVLinea = 0
                Me.dbcVLinea.Text = C_NINGUNA
                Me.dbcVLinea.Enabled = False
                mblnFueraChange = False
        End Select
    End Sub

    Private Sub dbcJFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Leave
        'gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " and DescFamilia LIKE '" & Trim(dbcJFamilia) & "%' ORDER BY DescFamilia"
        'ModDCombo.DCLostFocus dbcJFamilia, gStrSql, mintJFamilia

        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " and DescFamilia LIKE '" & Trim(dbcJFamilia.Text) & "%' ORDER BY DescFamilia"
        Aux = mintJFamilia
        mintJFamilia = 0
        If Trim(Me.dbcJFamilia.Text) <> Trim(C_TODAS) Or Trim(Me.dbcJFamilia.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcJFamilia), gStrSql, mintJFamilia)
        End If
        If Aux <> mintJFamilia Then
            If mintJFamilia = 0 Then
                mblnFueraChange = True
                Me.dbcJFamilia.Text = C_TODAS
                Me.dbcJFamilia.Enabled = True
                mintJLinea = 0
                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = False
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintJLinea = 0
                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = True
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
                Me.dbcJLinea.Focus()
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcJFamilia.Text) = "" Then Me.dbcJFamilia.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub dbcJLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.Leave
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ") and DescLinea LIKE '" & Trim(dbcJLinea.Text) & "%' ORDER BY DescLinea"
        Aux = mintJLinea
        mintJLinea = 0
        If Trim(Me.dbcJLinea.Text) <> Trim(C_TODAS) Or Trim(Me.dbcJLinea.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcJLinea), gStrSql, mintJLinea)
        End If
        If Aux <> mintJLinea Then
            If mintJLinea = 0 Then
                mblnFueraChange = True
                Me.dbcJLinea.Text = C_TODAS
                Me.dbcJLinea.Enabled = True
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintJSubLinea = 0
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = True
                mblnFueraChange = False
                Me.dbcJSubLinea.Focus()
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcJLinea.Text) = "" Then Me.dbcJLinea.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub dbcJSubLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Leave

        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ")  And (CodLinea = " & mintJLinea & ") and DescSubLinea LIKE '" & Trim(dbcJSubLinea.Text) & "%' ORDER BY DescSubLinea"
        Aux = mintJSubLinea
        mintJSubLinea = 0
        If Trim(Me.dbcJSubLinea.Text) <> Trim(C_TODAS) Or Trim(Me.dbcJSubLinea.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcJSubLinea), gStrSql, mintJSubLinea)
        End If
        If Aux <> mintJSubLinea Then
            If mintJSubLinea = 0 Then
                mblnFueraChange = True
                Me.dbcJSubLinea.Text = C_TODAS
                Me.dbcJSubLinea.Enabled = True
                mblnFueraChange = False
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcJSubLinea.Text) = "" Then Me.dbcJSubLinea.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub dbcOrigen1_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen1.CursorChanged
        On Error GoTo MError
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me.dbcOrigen1.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, (Me.dbcOrigen1))
        intCodOrigen = -1
        mblnFueraChange = True
        txtCodOrigen.Text = ""
        mblnFueraChange = False
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcOrigen1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen1.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen ORDER BY CodAlmacenOrigen"
        ModDCombo.DCGotFocus(gStrSql, dbcOrigen1)
    End Sub

    Private Sub dbcOrigen1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcOrigen1.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.txtCodOrigen.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcOrigen1_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcOrigen1.KeyUp
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Up Or eventArgs.KeyCode = System.Windows.Forms.Keys.Down Then
            PonerCodigoOrigen()
            Exit Sub
        End If
    End Sub

    Private Sub dbcOrigen1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen1.Leave
        Dim I As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen Where descAlmacenOrigen LIKE '" & Trim(Me.dbcOrigen1.Text) & "%'"
        intCodOrigen = -1
        ModDCombo.DCLostFocus((Me.dbcOrigen1), gStrSql, intCodOrigen)
        mblnFueraChange = True
        If intCodOrigen = -1 Or Trim(dbcOrigen1.Text) = "" Then
            txtCodOrigen.Text = ""
        Else
            txtCodOrigen.Text = CStr(intCodOrigen)
        End If
        mblnFueraChange = False
    End Sub

    Private Sub dbcOrigen1_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcOrigen1.MouseUp
        PonerCodigoOrigen()
    End Sub

    Private Sub dbcRMarca_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcRMarca.Name Then Exit Sub
        gStrSql = "SELECT CodMarca , DescMarca =ltrim(rtrim(DescMarca))  From CatMarcas Where CodGRupo = " & gCODRELOJERIA & " and DescMarca LIKE '" & Trim(dbcRMarca.Text) & "%' ORDER BY DescMarca"
        ModDCombo.DCChange(gStrSql, tecla)
        '    LimpiaDatosMarca
    End Sub

    Private Sub dbcRMarca_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.Enter
        Pon_Tool()
        gStrSql = "SELECT CodMarca , DescMarca =ltrim(rtrim(DescMarca))  From CatMarcas Where CodGRupo = " & gCODRELOJERIA & " ORDER BY DescMarca"
        ModDCombo.DCGotFocus(gStrSql, dbcRMarca)
    End Sub

    Private Sub dbcRMarca_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcRMarca.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkRelojeria.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcRMarca_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRMarca.Leave
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodMarca , DescMarca =ltrim(rtrim(DescMarca))  From CatMarcas Where CodGRupo = " & gCODRELOJERIA & " and DescMarca LIKE '" & Trim(dbcRMarca.Text) & "%' ORDER BY DescMarca"
        Aux = mintRMarca
        mintRMarca = 0
        If Trim(Me.dbcRMarca.Text) <> Trim(C_TODAS) Or Trim(Me.dbcRMarca.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcRMarca), gStrSql, mintRMarca)
        End If

        If Aux <> mintRMarca Then
            If mintRMarca = 0 Then
                mblnFueraChange = True
                Me.dbcRMarca.Text = C_TODAS
                Me.dbcRMarca.Enabled = True
                mintRModelo = 0
                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintRModelo = 0
                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = True
                mblnFueraChange = False
                Me.dbcRModelo.Focus()
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcRMarca.Text) = "" Then Me.dbcRMarca.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    '''Relojeria --Modelos
    Private Sub dbcRmodelo_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcRModelo.Name Then Exit Sub
        gStrSql = "SELECT Codmodelo , Descmodelo =ltrim(rtrim(Descmodelo))  From Catmodelos Where CodGRupo = " & gCODRELOJERIA & " And CodMarca = " & mintRMarca & " and Descmodelo LIKE '" & Trim(dbcRModelo.Text) & "%' ORDER BY Descmodelo"
        ModDCombo.DCChange(gStrSql, tecla)
        '    LimpiaDatosPrecioYDescuento
    End Sub

    Private Sub dbcRmodelo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.Enter
        Pon_Tool()
        gStrSql = "SELECT Codmodelo , Descmodelo =ltrim(rtrim(Descmodelo))  From Catmodelos Where CodGRupo = " & gCODRELOJERIA & " And CodMarca = " & mintRMarca & " ORDER BY Descmodelo"
        ModDCombo.DCGotFocus(gStrSql, dbcRModelo)
    End Sub

    Private Sub dbcRmodelo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcRModelo.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcRMarca.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcRModelo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.Leave
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT Codmodelo , Descmodelo =ltrim(rtrim(Descmodelo))  From Catmodelos Where CodGRupo = " & gCODRELOJERIA & " And CodMarca = " & mintRMarca & " and Descmodelo LIKE '" & Trim(dbcRModelo.Text) & "%' ORDER BY Descmodelo"
        Aux = mintRModelo
        mintRModelo = 0
        If Trim(Me.dbcRModelo.Text) <> Trim(C_TODOS) Or Trim(Me.dbcRModelo.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcRModelo), gStrSql, mintRModelo)
        End If
        If Aux <> mintRModelo Then
            If mintRModelo = 0 Then
                mblnFueraChange = True
                Me.dbcRModelo.Text = C_TODOS
                Me.dbcRModelo.Enabled = True
                mblnFueraChange = False
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcRModelo.Text) = "" Then Me.dbcRModelo.Text = C_TODOS
        mblnFueraChange = False
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged

        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        mblnFueraChange = True
        txtCodSucursal.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        '    If Screen.ActiveForm.ActiveControl.Name <> dbcSucursales.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen where TipoAlmacen ='P'ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            ModEstandar.RetrocederTab(Me)
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursales_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyUp
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Up Or eventArgs.KeyCode = System.Windows.Forms.Keys.Down Then
            PonerCodigoSucursal()
            Exit Sub
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        mblnFueraChange = True
        If intCodSucursal = 0 Then
            txtCodSucursal.Text = ""
        Else
            txtCodSucursal.Text = VB6.Format(intCodSucursal, "000")
        End If
        mblnFueraChange = False
    End Sub
    Private Sub dbcSucursales_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSucursales.MouseUp
        PonerCodigoSucursal()
    End Sub

    Private Sub dtpFechaCorte_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dtpFechaCorte.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
        End If
    End Sub

    Private Sub frmRptExistenciasyCostos_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmRptExistenciasyCostos_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub Form_Initialize_Renamed()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frmRptExistenciasyCostos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmRptExistenciasyCostos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub Nuevo()
        dtpFechaCorte.Value = Today
        Me.chkJoyeria.CheckState = System.Windows.Forms.CheckState.Checked
        chkJoyeria_CheckStateChanged(chkJoyeria, New System.EventArgs())

        Me.chkRelojeria.CheckState = System.Windows.Forms.CheckState.Checked
        chkRelojeria_CheckStateChanged(chkRelojeria, New System.EventArgs())

        Me.chkVarios.CheckState = System.Windows.Forms.CheckState.Checked
        chkVarios_CheckStateChanged(chkVarios, New System.EventArgs())

        intCodSucursal = 0
        mintJFamilia = 0
        mintJLinea = 0
        mintJSubLinea = 0
        mintRMarca = 0
        mintRModelo = 0
        mintVFamilia = 0
        mintVLinea = 0
        txtRangoDesde.Text = ""
        txtRangoHasta.Text = ""
        chkRangoTodos.CheckState = System.Windows.Forms.CheckState.Checked
        chkRangoSinExistencia.CheckState = System.Windows.Forms.CheckState.Checked
        mblnFueraChange = True
        txtCodOrigen.Text = ""
        dbcOrigen1.Text = ""
        mblnFueraChange = False
        chkMostrarAparatdos.CheckState = System.Windows.Forms.CheckState.Unchecked
        optOrdArticulo.Checked = True
        chkMostrarAparatdos.CheckState = System.Windows.Forms.CheckState.Checked
        optAlCosto.Checked = True

        optxSuc.Checked = True
        Label1(0).Enabled = True
        Label2.Enabled = True
        txtCodSucursal.Enabled = True
        dbcSucursales.Enabled = True
        txtCodOrigen.Enabled = True
        dbcOrigen1.Enabled = True
        chkJoyeria.Enabled = True
        dbcJFamilia.Enabled = True
        dbcJLinea.Enabled = True
        dbcJSubLinea.Enabled = True
        chkRelojeria.Enabled = True
        dbcRMarca.Enabled = True
        dbcRModelo.Enabled = True
        chkVarios.Enabled = True
        dbcVFamilia.Enabled = True
        dbcVLinea.Enabled = True
        fraOrdenamiento.Enabled = True
        optOrdArticulo.Enabled = True
        optOrdExistencia.Enabled = True
        cmbOrdArticulo.Enabled = True
        cmbOrdExistencia.Enabled = True

    End Sub

    Private Sub frmRptExistenciasyCostos_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        dtpFechaCorte.MinDate = C_FECHAINICIAL
        dtpFechaCorte.MaxDate = C_FECHAFINAL
        Nuevo()
    End Sub

    Private Sub frmRptExistenciasyCostos_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmRptExistenciasyCostos_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub Limpiar()
        Nuevo()
        dtpFechaCorte.Focus()
    End Sub

    Private Sub dbcJFAmilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcJFamilia.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkJoyeria.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcJFAmilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcJFamilia.Name Then Exit Sub
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " and DescFamilia LIKE '" & Trim(dbcJFamilia.Text) & "%' ORDER BY DescFamilia"
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcjFAmilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " ORDER BY DescFamilia"
        ModDCombo.DCGotFocus(gStrSql, dbcJFamilia)
    End Sub

    Private Sub dbcJLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcJLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcJFamilia.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode

    End Sub

    Private Sub dbcJLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcJLinea.Name Then Exit Sub
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ") and DescLinea LIKE '" & Trim(dbcJLinea.Text) & "%' ORDER BY DescLinea"
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcJLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJLinea.Enter
        If mblnFueraChange = True Then Exit Sub

        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ")  ORDER BY DescLinea"
        ModDCombo.DCGotFocus(gStrSql, dbcJLinea)
    End Sub

    Private Sub dbcJSubLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcJSubLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcJLinea.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcJSubLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcJSubLinea.Name Then Exit Sub
        gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ")  And (CodLinea = " & mintJLinea & ") and DescSubLinea LIKE '" & Trim(dbcJSubLinea.Text) & "%' ORDER BY DescSubLinea"
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcJSubLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Enter
        Pon_Tool()
        gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ")  And (CodLinea = " & mintJLinea & ") ORDER BY DescSubLinea"
        ModDCombo.DCGotFocus(gStrSql, dbcJSubLinea)
    End Sub

    Private Sub dbcVFamilia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcVFamilia.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.chkVarios.Focus()
            eventSender.KeyCode = 0
        ElseIf eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
            '        AvanzarTab Me
            '        dbcVFamilia_LostFocus
            '        KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcVFamilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcVFamilia.Name Then Exit Sub
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " and DescFamilia LIKE '" & Trim(dbcVFamilia.Text) & "%' ORDER BY DescFamilia"
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcVFamilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " ORDER BY DescFamilia"
        ModDCombo.DCGotFocus(gStrSql, dbcVFamilia)
    End Sub

    Private Sub dbcVFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " and DescFamilia LIKE '" & Trim(dbcVFamilia.Text) & "%' ORDER BY DescFamilia"
        Aux = mintVFamilia
        mintVFamilia = 0
        If Trim(Me.dbcVFamilia.Text) <> Trim(C_TODAS) Or Trim(Me.dbcVFamilia.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcVFamilia), gStrSql, mintVFamilia)
        End If

        If Aux <> mintVFamilia Then
            If mintVFamilia = 0 Then
                mblnFueraChange = True
                Me.dbcVFamilia.Text = C_TODAS
                Me.dbcVFamilia.Enabled = True
                mintVLinea = 0
                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = False
                mblnFueraChange = False
            Else
                mblnFueraChange = True
                mintVLinea = 0
                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = True
                mblnFueraChange = False
                Me.dbcVLinea.Focus()
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcVFamilia.Text) = "" Then Me.dbcVFamilia.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub dbcVLinea_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcVLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcVFamilia.Focus()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcVLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcVLinea.Name Then Exit Sub
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & mintVFamilia & ") and DescLinea LIKE '" & Trim(dbcVLinea.Text) & "%' ORDER BY DescLinea"
        ModDCombo.DCChange(gStrSql, tecla)
    End Sub

    Private Sub dbcVLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.Enter
        If mblnFueraChange = True Then Exit Sub
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & mintVFamilia & ")  ORDER BY DescLinea"
        ModDCombo.DCGotFocus(gStrSql, dbcVLinea)
    End Sub

    Private Sub dbcVLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.Leave
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'End If
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & mintVFamilia & ") and DescLinea LIKE '" & Trim(dbcVLinea.Text) & "%' ORDER BY DescLinea"
        Aux = mintVLinea
        mintVLinea = 0
        If Trim(Me.dbcVLinea.Text) <> Trim(C_TODAS) Or Trim(Me.dbcVLinea.Text) = "" Then
            ModDCombo.DCLostFocus((Me.dbcVLinea), gStrSql, mintVLinea)
        End If
        If Aux <> mintVLinea Then
            If mintVLinea = 0 Then
                mblnFueraChange = True
                Me.dbcVLinea.Text = C_TODAS
                Me.dbcVLinea.Enabled = True
                mblnFueraChange = False
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcVLinea.Text) = "" Then Me.dbcVLinea.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub optAlCosto_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAlCosto.CheckedChanged
        If eventSender.Checked Then
            If optAlCosto.Checked = True Then
                chkIncluirIVA.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkIncluirIVA.Enabled = False
                chkCostoAdicional.Enabled = False
                chkCostoFactura.Enabled = False
                chkCostoIndirecto.Enabled = False
                chkCostoFactura.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkCostoIndirecto.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkCostoAdicional.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
        End If
    End Sub
    Private Sub optOrdArticulo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optOrdArticulo.CheckedChanged
        If eventSender.Checked Then
            If optOrdArticulo.Checked = True Then
                cmbOrdArticulo.Enabled = True
                cmbOrdArticulo.SelectedIndex = 0
                cmbOrdExistencia.Text = ""
                cmbOrdExistencia.Enabled = False
            End If
        End If
    End Sub

    Private Sub optOrdExistencia_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optOrdExistencia.CheckedChanged
        If eventSender.Checked Then
            If optOrdExistencia.Checked = True Then
                cmbOrdExistencia.Enabled = True
                cmbOrdExistencia.SelectedIndex = 0
                cmbOrdArticulo.Text = ""
                cmbOrdArticulo.Enabled = False
            End If
        End If
    End Sub

    Private Sub optPrecioPublico_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPrecioPublico.CheckedChanged
        If eventSender.Checked Then
            If optPrecioPublico.Checked = True Then
                chkIncluirIVA.Enabled = True
                chkCostoAdicional.Enabled = False
                chkCostoFactura.Enabled = False
                chkCostoIndirecto.Enabled = False
                chkCostoFactura.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkCostoIndirecto.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkCostoAdicional.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
        End If
    End Sub

    Private Sub optTotal_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTotal.CheckedChanged
        If eventSender.Checked Then
            If optTotal.Checked = True Then
                Label1(0).Enabled = False
                Label2.Enabled = False
                txtCodSucursal.Enabled = False
                dbcSucursales.Enabled = False
                txtCodOrigen.Enabled = False
                dbcOrigen1.Enabled = False
                chkJoyeria.Enabled = False
                dbcJFamilia.Enabled = False
                dbcJLinea.Enabled = False
                dbcJSubLinea.Enabled = False
                chkRelojeria.Enabled = False
                dbcRMarca.Enabled = False
                dbcRModelo.Enabled = False
                chkVarios.Enabled = False
                dbcVFamilia.Enabled = False
                dbcVLinea.Enabled = False
                fraOrdenamiento.Enabled = False
                optOrdArticulo.Enabled = False
                optOrdExistencia.Enabled = False
                cmbOrdArticulo.Enabled = False
                cmbOrdExistencia.Enabled = False
            End If
        End If
    End Sub

    Private Sub optUltimoCostoPesos_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUltimoCostoPesos.CheckedChanged
        If eventSender.Checked Then
            chkCostoAdicional.Enabled = True
            chkCostoFactura.Enabled = True
            chkCostoIndirecto.Enabled = True
            chkIncluirIVA.Enabled = False
            chkCostoFactura.CheckState = System.Windows.Forms.CheckState.Checked
            chkCostoIndirecto.CheckState = System.Windows.Forms.CheckState.Checked
        End If
    End Sub

    Private Sub optxSuc_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optxSuc.CheckedChanged
        If eventSender.Checked Then
            If optxSuc.Checked = True Then
                Label1(0).Enabled = True
                Label2.Enabled = True
                txtCodSucursal.Enabled = True
                dbcSucursales.Enabled = True
                txtCodOrigen.Enabled = True
                dbcOrigen1.Enabled = True
                chkJoyeria.Enabled = True
                dbcJFamilia.Enabled = True
                dbcJLinea.Enabled = True
                dbcJSubLinea.Enabled = True
                chkRelojeria.Enabled = True
                dbcRMarca.Enabled = True
                dbcRModelo.Enabled = True
                chkVarios.Enabled = True
                dbcVFamilia.Enabled = True
                dbcVLinea.Enabled = True
                fraOrdenamiento.Enabled = True
                optOrdArticulo.Enabled = True
                optOrdExistencia.Enabled = True
                cmbOrdArticulo.Enabled = True
                cmbOrdExistencia.Enabled = True
            End If
        End If
    End Sub

    Private Sub txtCodOrigen_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodOrigen.TextChanged
        If mblnFueraChange = True Then Exit Sub
        mblnFueraChange = True
        dbcOrigen1.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub txtCodOrigen_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodOrigen.Enter
        SelTextoTxt(txtCodOrigen)
    End Sub


    Private Sub txtCodOrigen_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCodOrigen.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCodOrigen_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodOrigen.Leave
        LlenaDatosOrigen()
    End Sub

    Private Sub txtCodSucursal_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.TextChanged
        If mblnFueraChange = True Then Exit Sub
        mblnFueraChange = True
        dbcSucursales.Text = ""
        mblnFueraChange = False
    End Sub

    Private Sub txtCodSucursal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Enter
        SelTextoTxt(txtCodSucursal)
    End Sub

    Private Sub txtCodSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCodSucursal.Leave
        LlenaDatosSucursal()
    End Sub

    Sub LlenaDatosSucursal()
        If CDbl(Numerico(Trim(txtCodSucursal.Text))) = 0 Then Exit Sub
        gStrSql = "SELECT      Ltrim(Rtrim(DescAlmacen)) as DescAlmacen From dbo.CatAlmacen Where CodAlmacen =" & Numerico(txtCodSucursal.Text) & "  And TipoAlmacen = 'P'"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            mblnFueraChange = True
            dbcSucursales.Text = RsGral.Fields("DescAlmacen").Value
            mblnFueraChange = False
        Else
            MsgBox("Código de Sucursal no existe." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            txtCodSucursal.Text = ""
            dbcSucursales.Focus()
            Exit Sub
        End If
    End Sub

    Sub LlenaDatosOrigen()
        On Error GoTo Merr
        If Trim(txtCodOrigen.Text) = "" Then Exit Sub
        gStrSql = "SELECT  CodAlmacenOrigen, Ltrim(Rtrim(DescAlmacenOrigen)) as DescALmacenOrigen From dbo.CatOrigen Where CodAlmacenOrigen = " & (txtCodOrigen).Text
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        If RsGral.RecordCount > 0 Then
            mblnFueraChange = True
            dbcOrigen1.Text = RsGral.Fields("DescAlmacenorigen").Value
            mblnFueraChange = False
        Else
            MsgBox("Código de Origen no existe." & vbNewLine & "Verifique por favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            txtCodOrigen.Text = ""
            dbcOrigen1.Focus()
        End If
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Public Function DevuelveQuery() As String
        On Error GoTo Merr
        Dim I As Integer
        Dim cSELECT As String
        Dim cFROM As String
        Dim cWHERE As String
        Dim cGROUPBY As String
        Dim cORDERBY As String
        Dim rsLocal As ADODB.Recordset
        Dim cMSG As String
        Dim lRedondeo As Integer
        Dim cHAVING As String
        Dim cPorcIva As String

        Dim nJOYERIA As Integer
        Dim nRELOJERIA As Integer
        Dim nVARIOS As Integer

        'Obtener los códigos que va a tomar en cuenta en la consulta; estos códigos se enviarán como parámetros al
        'procedimiento almacenado que recopilará los datos

        'reporte por sucursal
        If optxSuc.Checked Then

            nJOYERIA = Me.chkJoyeria.CheckState
            nRELOJERIA = Me.chkRelojeria.CheckState
            nVARIOS = Me.chkVarios.CheckState

            If nJOYERIA = 0 And nRELOJERIA = 0 And nVARIOS = 0 Then
                MsgBox("Debe elegir, por lo menos, un grupo con el cual generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                Exit Function
            End If

            cWHERE = " Having  "
            cSELECT = ""
            cGROUPBY = ""
            cORDERBY = ""

            Select Case True
                Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS > 0
                    'Todos los grupos
                    cWHERE = cWHERE & " A.CodGrupo In (" & gCODJOYERIA & ", " & gCODRELOJERIA & ", " & gCODVARIOS & ") "
                    Select Case True
                        Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                            ' Todos
                            cWHERE = cWHERE & " and ((A.CodFamilia <> " & 0 & " )"
                        Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                            cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & ")"
                        Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                            cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " )"
                        Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                            cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & ")"
                    End Select
                    Select Case True
                        Case mintRMarca <= 0 And mintRModelo <= 0
                            'Todos
                            cWHERE = cWHERE & " or (A.CodMarca <> " & 0 & ")"
                        Case mintRMarca > 0 And mintRModelo <= 0
                            cWHERE = cWHERE & " or (A.CodMarca = " & mintRMarca & ")"
                        Case mintRMarca > 0 And mintRModelo > 0
                            cWHERE = cWHERE & " or (A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & ")"
                    End Select
                    Select Case True
                        Case mintVFamilia <= 0 And mintVLinea <= 0
                            'Todos
                            cWHERE = cWHERE & " or (A.CodFamilia <> 0 and A.CodSubLinea is NULL))"
                        Case mintVFamilia > 0 And mintVLinea <= 0
                            cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL))"
                        Case mintVFamilia > 0 And mintVLinea > 0
                            cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL))"
                    End Select
                Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS <= 0
                    'Joyeria-Relojeria
                    cWHERE = cWHERE & " A.CodGrupo <> " & gCODVARIOS
                    Select Case True
                        Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                            ' Todos
                            cWHERE = cWHERE & " and ((A.CodFamilia <> " & 0 & " )"
                        Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                            cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " )"
                        Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                            cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " )"
                        Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                            cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & ")"
                    End Select
                    Select Case True
                        Case mintRMarca <= 0 And mintRModelo <= 0
                            'Todos
                            cWHERE = cWHERE & " or (A.CodMarca <> " & 0 & "))"
                        Case mintRMarca > 0 And mintRModelo <= 0
                            cWHERE = cWHERE & " or (A.CodMarca = " & mintRMarca & "))"
                        Case mintRMarca > 0 And mintRModelo > 0
                            cWHERE = cWHERE & " or (CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & "))"
                    End Select
                Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS > 0
                    'Joyeria-Varios
                    cWHERE = cWHERE & " A.CodGrupo <> " & gCODRELOJERIA
                    Select Case True
                        Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                            ' Todos
                            cWHERE = cWHERE & " and ((A.CodFamilia <> " & 0 & " )"
                        Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                            cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " )"
                        Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                            cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " )"
                        Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                            cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & ")"
                    End Select
                    Select Case True
                        Case mintVFamilia <= 0 And mintVLinea <= 0
                            'Todos
                            cWHERE = cWHERE & " or (A.CodFamilia <> 0) and A.CodSubLinea is NULL)"
                        Case mintVFamilia > 0 And mintVLinea <= 0
                            cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL))"
                        Case mintVFamilia > 0 And mintVLinea > 0
                            cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL))"
                    End Select
                Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS <= 0
                    'Joyeria
                    cWHERE = cWHERE & " A.CodGrupo = " & gCODJOYERIA
                    Select Case True
                        Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                            ' Todos
                            cWHERE = cWHERE & " and A.CodFamilia <> " & 0
                        Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                            cWHERE = cWHERE & " and A.CodFamilia = " & mintJFamilia
                        Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                            cWHERE = cWHERE & " and A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea
                        Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                            cWHERE = cWHERE & " and A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea
                    End Select
                Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS > 0
                    'Relojeria-Varios
                    cWHERE = cWHERE & " A.CodGrupo <> " & gCODJOYERIA
                    Select Case True
                        Case mintRMarca <= 0 And mintRModelo <= 0
                            'Todos
                            cWHERE = cWHERE & " and ((A.CodMarca <> " & 0 & ")"
                        Case mintRMarca > 0 And mintRModelo <= 0
                            cWHERE = cWHERE & " and ((A.CodMarca = " & mintRMarca & ")"
                        Case mintRMarca > 0 And mintRModelo > 0
                            cWHERE = cWHERE & " and ((A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & ")"
                    End Select
                    Select Case True
                        Case mintVFamilia <= 0 And mintVLinea <= 0
                            'Todos
                            cWHERE = cWHERE & " or (A.CodFamilia <> 0) and A.CodSubLinea is NULL)"
                        Case mintVFamilia > 0 And mintVLinea <= 0
                            cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL))"
                        Case mintVFamilia > 0 And mintVLinea > 0
                            cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL))"
                    End Select
                Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS <= 0
                    'Relojeria
                    cWHERE = cWHERE & " A.CodGrupo = " & gCODRELOJERIA
                    Select Case True
                        Case mintRMarca <= 0 And mintRModelo <= 0
                            'Todos
                            cWHERE = cWHERE & " and A.CodMarca <> " & 0
                        Case mintRMarca > 0 And mintRModelo <= 0
                            cWHERE = cWHERE & " and A.CodMarca = " & mintRMarca
                        Case mintRMarca > 0 And mintRModelo > 0
                            cWHERE = cWHERE & " and A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo
                    End Select
                Case nJOYERIA <= 0 And nRELOJERIA <= 0 And nVARIOS > 0
                    'Varios
                    cWHERE = cWHERE & " A.CodGrupo = " & gCODVARIOS
                    Select Case True
                        Case mintVFamilia <= 0 And mintVLinea <= 0
                            'Todos
                            cWHERE = cWHERE & " and A.CodFamilia <> 0 and A.CodSubLinea is NULL "
                        Case mintVFamilia > 0 And mintVLinea <= 0
                            cWHERE = cWHERE & " and A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL "
                        Case mintVFamilia > 0 And mintVLinea > 0
                            cWHERE = cWHERE & " and A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL "
                    End Select
            End Select

            If chkRangoSinExistencia.CheckState = System.Windows.Forms.CheckState.Unchecked And chkRangoTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
                cWHERE = cWHERE & " And (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) <> 0 "
            End If
            If chkRangoTodos.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                cWHERE = cWHERE & " And (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados))  BetWeen " & Numerico(txtRangoDesde.Text) & " and " & Numerico(txtRangoHasta.Text)
            End If

            If dbcOrigen1.Text <> "" Then
                cWHERE = cWHERE & " And A.CodAlmacenOrigen = " & Numerico(txtCodOrigen.Text)
            End If

            'Agregar el Almacen al WHERE
            cWHERE = cWHERE & " And I.CodAlmacen= " & Numerico(txtCodSucursal.Text) & "  "

            Select Case True
                Case optOrdArticulo.Checked = True And cmbOrdArticulo.Text = "Codigo"
                    cORDERBY = " Order By I.CodArticulo "
                Case optOrdArticulo.Checked = True And cmbOrdArticulo.Text = "Descripcion"
                    cORDERBY = " Order By A.DescArticulo "
                Case optOrdExistencia.Checked = True And cmbOrdExistencia.Text = "Ascendente"
                    cORDERBY = " Order By Existencia Asc"
                Case optOrdExistencia.Checked = True And cmbOrdExistencia.Text = "Descendente"
                    cORDERBY = " Order By Existencia Desc"
            End Select

            'Armar el String  para el Select
            If mintJFamilia > 0 Or mintVFamilia > 0 Then
                cSELECT = cSELECT & " ISNULL(A.CodFamilia, 0) AS CodFamilia, ISNULL(Fa.DescFamilia, '') AS DescFamilia,  "
                cGROUPBY = cGROUPBY & " CodFamilia, "
            End If
            If mintJLinea > 0 Or mintVLinea > 0 Then
                cSELECT = cSELECT & " ISNULL(A.CodLinea, 0) AS CodLinea, ISNULL(Li.DescLinea, '') AS DescLinea,  "
                cGROUPBY = cGROUPBY & " CodLinea, "
            End If
            If mintJSubLinea > 0 Then
                cSELECT = cSELECT & " ISNULL(A.CodSubLinea,0) AS CodSublinea, ISNULL(Su.DescSubLinea, '') AS DescSubLinea,  "
                cGROUPBY = cGROUPBY & " CodSubLinea, "
            End If
            If mintRMarca > 0 Then
                cSELECT = cSELECT & " ISNULL(A.CodMarca, 0) AS codmarca, ISNULL(Ma.DescMarca, '') AS DescMarca,  "
                cGROUPBY = cGROUPBY & " CodMarca, "
            End If
            If mintRModelo > 0 Then
                cSELECT = cSELECT & " ISNULL(A.CodModelo, 0) AS codmodelo, ISNULL(Mo.DescModelo, '') AS DescModelo,  "
                cGROUPBY = cGROUPBY & " CodModelo, "
            End If

            Dim fecha As String = AgregarHoraAFecha(dtpFechaCorte.Value)
            DevuelveQuery = "FROM dbo.Inventario I INNER JOIN " & "dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " & "dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " & "dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen LEFT OUTER JOIN " & "dbo.CatGrupos G ON A.CodGrupo = G.CodGrupo LEFT OUTER JOIN " & "dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia LEFT OUTER JOIN " & "dbo.CatLineas Li ON A.CodGrupo = Li.CodGrupo AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea LEFT OUTER JOIN " & "dbo.CatSubLineas Su ON A.CodGrupo = Su.CodGrupo AND A.CodFamilia = Su.CodFamilia AND A.CodLinea = Su.CodLinea AND " & "A.CodSubLinea = Su.CodSubLinea LEFT OUTER JOIN " & "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca LEFT OUTER JOIN " & "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND Mo.CodMarca = A.CodMarca AND A.CodModelo = Mo.CodModelo CROSS JOIN " & "dbo.ConfiguracionGeneral " & "WHERE     (I.FechaMovto <= '" & fecha & "') " & "GROUP BY I.CodArticulo, A.CodArticulo, A.DescArticulo, A.CodGrupo, A.CodUnidad, " & "O.DescAlmacenOrigen, Al.DescAlmacen, I.CodAlmacen, A.CodAlmacenOrigen, A.CostoReal, " & "dbo.ConfiguracionGeneral.NombreEmp, Al.DescAlmacen, A.DescArticulo, " & "G.DescGrupo ,A.CodFamilia,A.CodSubLinea,A.CodMarca,A.CodSubLinea, Fa.DescFamilia, " & "A.CodLinea,A.CodModelo,A.CodLinea , Li.DescLinea, Su.DescSubLinea, Ma.DescMarca,Mo.DescModelo, A.PrecioPubDolar,dbo.ConfiguracionGeneral.TasaIva  ," & "A.CostoFacturaPesos , A.CostoAdicionalPesos ,A.CostoIndirectoPesos ,PesosFijos " & cWHERE & " " & cORDERBY & "  "

            lRedondeo = 0
            If optAlCosto.Checked = True Then
                cSELECT = " SELECT     I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Convert(Integer, Convert(Integer, Round(A.CostoReal," & lRedondeo & "))) AS CostoUnitario, (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * (Convert(Integer, Round(A.CostoReal," & lRedondeo & "))) as Importe, " & "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
            ElseIf optPrecioPublico.Checked = True Then
                If chkIncluirIVA.CheckState = System.Windows.Forms.CheckState.Checked Then
                    cSELECT = " SELECT I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Case PesosFijos When 0 then Convert(Integer, Round(A.PrecioPubDolar," & lRedondeo & ")) When 1 Then Convert(Integer, Round(A.PrecioPubDolar/" & gcurCorpoTIPOCAMBIODOLAR & "," & lRedondeo & "))  End AS CostoUnitario,  " & "(SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * Case PesosFijos When 0 then Convert(Integer, Round(A.PrecioPubDolar," & lRedondeo & ")) When 1 Then Convert(Integer, Round(A.PrecioPubDolar/" & gcurCorpoTIPOCAMBIODOLAR & "," & lRedondeo & ")) End AS Importe, " & "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
                Else
                    cSELECT = " SELECT     I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Case PesosFijos When 0 then Convert(Integer, Round(A.PrecioPubDolar / (1 + dbo.ConfiguracionGeneral.TasaIva / 100) ," & lRedondeo & ")) When 1 Then Convert(Integer, Round((A.PrecioPubDolar  / (1 + dbo.ConfiguracionGeneral.TasaIva / 100) )/" & gcurCorpoTIPOCAMBIODOLAR & "," & lRedondeo & "))  End AS CostoUnitario, " & "(SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * Convert(Integer, Round(Case PesosFijos When 0 then A.PrecioPubDolar / (1 + dbo.ConfiguracionGeneral.TasaIva / 100) When 1 Then (A.PrecioPubDolar  / (1 + dbo.ConfiguracionGeneral.TasaIva / 100) )/" & gcurCorpoTIPOCAMBIODOLAR & " End ," & lRedondeo & "))  AS Importe, " & "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
                End If
            Else
                Select Case True
                    Case chkCostoFactura.CheckState = 1 And chkCostoAdicional.CheckState = 1 And chkCostoIndirecto.CheckState = 1
                        cSELECT = "SELECT I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Convert(Integer, Round(A.CostoFacturaPesos+ A.CostoAdicionalPesos + A.CostoIndirectoPesos," & lRedondeo & ")) AS CostoUnitario, " & "(SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * Convert(Integer, Round(A.CostoFacturaPesos+ A.CostoAdicionalPesos + A.CostoIndirectoPesos," & lRedondeo & ")) AS Importe, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
                    Case chkCostoFactura.CheckState = 1 And chkCostoAdicional.CheckState = 1 And chkCostoIndirecto.CheckState = 0
                        cSELECT = " SELECT     I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Convert(Integer, Round(A.CostoFacturaPesos+ A.CostoAdicionalPesos, " & lRedondeo & ")) AS CostoUnitario, (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * Convert(Integer, Round(A.CostoFacturaPesos+ A.CostoAdicionalPesos, " & lRedondeo & ")) AS Importe, " & "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
                    Case chkCostoFactura.CheckState = 1 And chkCostoAdicional.CheckState = 0 And chkCostoIndirecto.CheckState = 1
                        cSELECT = " SELECT     I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Convert(Integer, Round(A.CostoFacturaPesos+  A.CostoIndirectoPesos," & lRedondeo & ")) AS CostoUnitario, (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * Convert(Integer, Round(A.CostoFacturaPesos +  A.CostoIndirectoPesos," & lRedondeo & ")) AS Importe, " & "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
                    Case chkCostoFactura.CheckState = 1 And chkCostoAdicional.CheckState = 0 And chkCostoIndirecto.CheckState = 0
                        cSELECT = " SELECT     I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Convert(Integer, Round(A.CostoFacturaPesos," & lRedondeo & ")) AS CostoUnitario, (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * Convert(Integer, Round(A.CostoFacturaPesos," & lRedondeo & ")) AS Importe, " & "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
                    Case chkCostoFactura.CheckState = 0 And chkCostoAdicional.CheckState = 1 And chkCostoIndirecto.CheckState = 0
                        cSELECT = " SELECT     I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Convert(Integer, Round(A.CostoAdicionalPesos," & lRedondeo & ")) AS CostoUnitario, (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * Convert(Integer, Round(A.CostoIndirectoPesos," & lRedondeo & ")) AS Importe, " & "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
                    Case chkCostoFactura.CheckState = 0 And chkCostoAdicional.CheckState = 0 And chkCostoIndirecto.CheckState = 1
                        cSELECT = " SELECT     I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Convert(Integer, Round(A.CostoIndirectoPesos," & lRedondeo & ")) AS CostoUnitario, " & "(SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * Convert(Integer, Round(A.CostoIndirectoPesos," & lRedondeo & ")) AS Importe, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
                    Case chkCostoFactura.CheckState = 0 And chkCostoAdicional.CheckState = 1 And chkCostoIndirecto.CheckState = 1
                        cSELECT = " SELECT     I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " & "AS Existencia, SUM(I.Apartados) AS Apartados, Convert(Integer, Round(A.CostoAdicionalPesos + A.CostoIndirectoPesos," & lRedondeo & ")) AS CostoUnitario, (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) * Convert(Integer, Round(A.CostoAdicionalPesos + A.CostoIndirectoPesos," & lRedondeo & ")) AS Importe, " & "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " & "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
                    Case chkCostoFactura.CheckState = 0 And chkCostoAdicional.CheckState = 0 And chkCostoIndirecto.CheckState = 0
                        MsgBox("Debe elegir, por lo menos, un costo en pesos para poder generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                        DevuelveQuery = ""
                        Exit Function
                End Select

            End If
            DevuelveQuery = cSELECT & DevuelveQuery
        End If 'opcion x sucursal

        'opcion inv total
        If optTotal.Checked Then

            cHAVING = ""
            cSELECT = ""
            cORDERBY = " Order By CodGrupo, PNivel "
            cPorcIva = Trim(CStr(1 + (gcurCorpoTASAIVA / 100)))
            'utiliza el TC del dia para convertir precios de articulos en pesos a dolares
            'utiliza la tasa de iva actual para calcular las precios publicos sin iva

            If chkRangoSinExistencia.CheckState = System.Windows.Forms.CheckState.Unchecked And chkRangoTodos.CheckState = System.Windows.Forms.CheckState.Checked Then
                cHAVING = " Having Sum(Existencia) <> 0  "
            End If
            If chkRangoTodos.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                cHAVING = " Having sum(Existencia) Between " & CInt(Numerico(txtRangoDesde.Text)) & " and " & CInt(Numerico(txtRangoHasta.Text)) & "  "
            End If

            If optAlCosto.Checked Then
                cSELECT = "Select CodGrupo, DescGrupo, PNivel, LTRIM(RTRIM(NombreEmp)) as NombreEmp, sum(Existencia) as Existencia, sum(Apartados) as Apartados, sum(Costo) as CostoUnitario, sum((Existencia * Costo)) as Importe  From ( "
            ElseIf optPrecioPublico.Checked Then
                If chkIncluirIVA.CheckState = 1 Then 'CON IVA
                    cSELECT = "Select CodGrupo, DescGrupo, PNivel, LTRIM(RTRIM(NombreEmp)) as NombreEmp, sum(Existencia) as Existencia, sum(Apartados) as Apartados, sum(PrecioP) as CostoUnitario, sum((Existencia * PrecioP)) as Importe  From ( "
                Else 'SIN IVA
                    cSELECT = "Select CodGrupo, DescGrupo, PNivel, LTRIM(RTRIM(NombreEmp)) as NombreEmp, sum(Existencia) as Existencia, sum(Apartados) as Apartados, sum(PrecioP) as CostoUnitario, Convert(Integer, (Sum(Existencia * PrecioP))/" & cPorcIva & ") as Importe  From ( "
                End If
            ElseIf optUltimoCostoPesos.Checked Then
                'DEFINIR LOS COSTOS QUE PARTICIPAN EN EL IMPORTE
                Select Case True
                    Case chkCostoFactura.CheckState = 1 And chkCostoAdicional.CheckState = 1 And chkCostoIndirecto.CheckState = 1
                        cSELECT = "Select CodGrupo, DescGrupo, PNivel, LTRIM(RTRIM(NombreEmp)) as NombreEmp, sum(Existencia) as Existencia, sum(Apartados) as Apartados, sum(CostoFacturaPesos + CostoAdicionalPesos + CostoIndirectoPesos) as CostoUnitario, sum(Existencia * (CostoFacturaPesos + CostoAdicionalPesos + CostoIndirectoPesos)) as Importe  From ( "
                    Case chkCostoFactura.CheckState = 1 And chkCostoAdicional.CheckState = 1 And chkCostoIndirecto.CheckState = 0
                        cSELECT = "Select CodGrupo, DescGrupo, PNivel, LTRIM(RTRIM(NombreEmp)) as NombreEmp, sum(Existencia) as Existencia, sum(Apartados) as Apartados, sum(CostoFacturaPesos + CostoAdicionalPesos) as CostoUnitario, sum(Existencia * (CostoFacturaPesos + CostoAdicionalPesos)) as Importe  From ( "
                    Case chkCostoFactura.CheckState = 1 And chkCostoAdicional.CheckState = 0 And chkCostoIndirecto.CheckState = 1
                        cSELECT = "Select CodGrupo, DescGrupo, PNivel, LTRIM(RTRIM(NombreEmp)) as NombreEmp, sum(Existencia) as Existencia, sum(Apartados) as Apartados, sum(CostoFacturaPesos + CostoIndirectoPesos) as CostoUnitario, sum(Existencia * (CostoFacturaPesos + CostoIndirectoPesos)) as Importe  From ( "
                    Case chkCostoFactura.CheckState = 1 And chkCostoAdicional.CheckState = 0 And chkCostoIndirecto.CheckState = 0
                        cSELECT = "Select CodGrupo, DescGrupo, PNivel, LTRIM(RTRIM(NombreEmp)) as NombreEmp, sum(Existencia) as Existencia, sum(Apartados) as Apartados, sum(CostoFacturaPesos) as CostoUnitario, sum(Existencia * CostoFacturaPesos) as Importe  From ( "

                    Case chkCostoFactura.CheckState = 0 And chkCostoAdicional.CheckState = 0 And chkCostoIndirecto.CheckState = 0
                        MsgBox("Debe elegir, por lo menos, un costo en pesos para poder generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
                        DevuelveQuery = ""
                        Exit Function
                End Select
            End If

            'Se elimino el tipo de almacen para que tome tambien los almacenes de VExternos
            '15NOV2004
            'gStrSql = cSELECT &
            '"SELECT A.CodGrupo, G.DescGrupo, A.CodArticulo, CodigoArticuloProv, convert(char(1), OrigenAnt) + '-' + right('00000'+ltrim(rtrim(convert(char(5), CodigoAnt))),5) as CodigoAnt,  Sum((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) AS Existencia, sum(I.Apartados) AS Apartados, " &
            '"Max(Convert(Integer, Convert(Integer, Round(A.CostoReal,0)))) as Costo, Max(Convert(Integer, Convert(Integer, Round(Case When A.PesosFijos = 1 Then (PrecioPubDolar/" & gcurCorpoTIPOCAMBIODOLAR & ") Else PrecioPubDolar End,0)))) as PrecioP, ltrim(rtrim(Case When A.CodGrupo = 1 Then Li.DescLinea When A.CodGrupo = 2 Then Ma.DescMarca When A.CodGrupo = 3 Then Fa.DescFamilia Else '' End)) as PNivel, LTRIM(RTrim(CG.NombreEmp)) As NombreEmp " &
            '"FROM  dbo.Inventario I INNER JOIN dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN dbo.CatAlmacen Al     ON I.CodAlmacen = Al.CodAlmacen /* And TipoAlmacen = 'P' */ LEFT OUTER JOIN dbo.CatGrupos G    ON A.CodGrupo = G.CodGrupo LEFT OUTER JOIN dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia LEFT OUTER JOIN dbo.CatLineas Li   ON A.CodGrupo = Li.CodGrupo AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea LEFT OUTER JOIN dbo.CatMarcas Ma   ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca CROSS JOIN dbo.ConfiguracionGeneral CG " &
            '"WHERE (I.FechaMovto <= '" & Format(dtpFechaCorte, C_FORMATFECHAGUARDAR) & "') /* And Al.CodAlmacen in (1) */ And  ( (A.CodGrupo = 1 And A.CodFamilia <> 0 ) or (A.CodGrupo = 2 And A.CodMarca <> 0) or (A.CodGrupo = 3 And A.CodFamilia <> 0 ) ) " &
            '"Group By A.CodGrupo, G.DescGrupo, A.CodArticulo, A.DescArticulo, CodigoArticuloProv, convert(char(1), OrigenAnt) + '-' + right('00000'+ltrim(rtrim(convert(char(5), CodigoAnt))),5), ltrim(rtrim(Case When A.CodGrupo = 1 Then Li.DescLinea When A.CodGrupo = 2 Then Ma.DescMarca When A.CodGrupo = 3 Then Fa.DescFamilia Else '' End)), LTRIM(RTRIM(CG.NombreEmp)) " &
            '") as InvTotal Group By CodGrupo, DescGrupo, PNivel, LTRIM(RTRIM(NombreEmp)) " & cHAVING & cORDERBY

            '     /* And TipoAlmacen = 'P' */
            '     /* And Al.CodAlmacen in (1) */   - eliminado de la anterior modific
            Dim fechaCorte As String = AgregarHoraAFecha(dtpFechaCorte.Value)

            gStrSql = cSELECT & "SELECT A.CodGrupo, G.DescGrupo, A.CodArticulo, CodigoArticuloProv, 
convert(char(1), OrigenAnt) + '-' + right('00000'+ltrim(rtrim(convert(char(5), CodigoAnt))),5) as CodigoAnt, 
Sum((I.ExistenciaInicial + I.Entradas) - (I.Salidas + I.Apartados)) AS Existencia, sum(I.Apartados) AS Apartados, 
" & "Max(Convert(Integer, Convert(Integer, Round(A.CostoReal,0)))) as Costo, Max(Convert(Integer, Convert(Integer, Round(Case When A.PesosFijos = 1 Then (PrecioPubDolar/
" & gcurCorpoTIPOCAMBIODOLAR & ") Else PrecioPubDolar End,0)))) as PrecioP, ltrim(rtrim(Case When A.CodGrupo = 1 Then Li.DescLinea When A.CodGrupo = 2 Then Ma.DescMarca When A.CodGrupo = 3 
Then Fa.DescFamilia Else '' End)) as PNivel, LTRIM(RTrim(CG.NombreEmp)) As NombreEmp, Max(Convert(Integer, Convert(Integer, Round(A.CostoFacturaPesos,0)))) as CostoFacturaPesos, 
" & "Max(Convert(Integer, Convert(Integer, Round(A.CostoAdicionalPesos,0)))) as CostoAdicionalPesos, Max(Convert(Integer, Convert(Integer, Round(A.CostoIndirectoPesos,0)))) as CostoIndirectoPesos 
" & "FROM  dbo.Inventario I INNER JOIN dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen LEFT OUTER JOIN dbo.CatGrupos G    
ON A.CodGrupo = G.CodGrupo LEFT OUTER JOIN dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia LEFT OUTER JOIN dbo.CatLineas Li   ON A.CodGrupo = Li.CodGrupo 
AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea LEFT OUTER JOIN dbo.CatMarcas Ma   ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca CROSS JOIN dbo.ConfiguracionGeneral CG 
" & "WHERE (I.FechaMovto <= '" & fechaCorte & "') And  ( (A.CodGrupo = 1 And A.CodFamilia <> 0 ) or (A.CodGrupo = 2 And A.CodMarca <> 0) or (A.CodGrupo = 3 And A.CodFamilia <> 0 ) ) 
" & "Group By A.CodGrupo, G.DescGrupo, A.CodArticulo, A.DescArticulo, CodigoArticuloProv, convert(char(1), OrigenAnt) + '-' + right('00000'+ltrim(rtrim(convert(char(5), CodigoAnt))),5), 
ltrim(rtrim(Case When A.CodGrupo = 1 Then Li.DescLinea When A.CodGrupo = 2 Then Ma.DescMarca When A.CodGrupo = 3 Then Fa.DescFamilia Else '' End)), LTRIM(RTRIM(CG.NombreEmp)) 
" & ") as InvTotal Group By CodGrupo, DescGrupo, PNivel, LTRIM(RTRIM(NombreEmp)) " & cHAVING & cORDERBY

            DevuelveQuery = gStrSql
        End If
        Exit Function

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function

    '    Public Function DevuelveQuery() As String
    '        On Local Error GoTo MErr
    '        Dim i As Long
    '        Dim cSELECT As String
    '        Dim cFROM As String
    '        Dim cWHERE As String
    '        Dim cGROUPBY As String
    '        Dim cORDERBY As String
    '        Dim rsLocal As ADODB.Recordset
    '        Dim cMSG As String

    '        Dim nJOYERIA As Long
    '        Dim nRELOJERIA As Long
    '        Dim nVARIOS As Long

    '        Obtener los códigos que va a tomar en cuenta en la consulta; estos códigos se enviarán como parámetros al
    '        procedimiento almacenado que recopilará los datos

    '        nJOYERIA = Me.chkJoyeria.Value
    '        nRELOJERIA = Me.chkRelojeria.Value
    '        nVARIOS = Me.chkVarios.Value

    '        If nJOYERIA = 0 And nRELOJERIA = 0 And nVARIOS = 0 Then
    '            MsgBox "Debe elegir, por lo menos, un grupo con el cual generar el reporte", vbOKOnly + vbInformation, gstrCorpoNOMBREEMPRESA
    '            Exit Function
    '        End If

    '        cWHERE = " Having  "
    '        cSELECT = ""
    '        cGROUPBY = ""
    '        cORDERBY = ""

    '        Select Case True
    '            Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS > 0
    '                Todos los grupos
    '                cWHERE = cWHERE & " A.CodGrupo In (" & gCODJOYERIA & ", " & gCODRELOJERIA & ", " & gCODVARIOS & ") "
    '                Select Case True
    '                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " and ((A.CodFamilia <> " & 0 & " and A.CodSubLinea is NOT NULL)"
    '                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodSubLinea is NOT NULL)"
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea is NOT NULL)"
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
    '                        cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & ")"
    '                End Select
    '                Select Case True
    '                    Case mintRMarca <= 0 And mintRModelo <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " or (A.CodMarca <> " & 0 & ")"
    '                    Case mintRMarca > 0 And mintRModelo <= 0
    '                        cWHERE = cWHERE & " or (A.CodMarca = " & mintRMarca & ")"
    '                    Case mintRMarca > 0 And mintRModelo > 0
    '                        cWHERE = cWHERE & " or (A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & ")"
    '                End Select
    '                Select Case True
    '                    Case mintVFamilia <= 0 And mintVLinea <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " or (A.CodFamilia <> 0 and A.CodSubLinea is NULL))"
    '                    Case mintVFamilia > 0 And mintVLinea <= 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL))"
    '                    Case mintVFamilia > 0 And mintVLinea > 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL))"
    '                End Select
    '            Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS <= 0
    '                Joyeria-Relojeria
    '                cWHERE = cWHERE & " A.CodGrupo <> " & gCODVARIOS
    '                Select Case True
    '                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " and ((A.CodFamilia <> " & 0 & " and A.CodSubLinea is NOT NULL)"
    '                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodSubLinea is NOT NULL)"
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea is NOT NULL)"
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
    '                        cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & ")"
    '                End Select
    '                Select Case True
    '                    Case mintRMarca <= 0 And mintRModelo <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " or (A.CodMarca <> " & 0 & "))"
    '                    Case mintRMarca > 0 And mintRModelo <= 0
    '                        cWHERE = cWHERE & " or (A.CodMarca = " & mintRMarca & "))"
    '                    Case mintRMarca > 0 And mintRModelo > 0
    '                        cWHERE = cWHERE & " or (CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & "))"
    '                End Select
    '            Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS > 0
    '                Joyeria-Varios
    '                cWHERE = cWHERE & " A.CodGrupo <> " & gCODRELOJERIA
    '                Select Case True
    '                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " and ((A.CodFamilia <> " & 0 & " and A.CodSubLinea is NOT NULL)"
    '                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodSubLinea is NOT NULL)"
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea is NOT NULL)"
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
    '                        cWHERE = cWHERE & " and ((A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & ")"
    '                End Select
    '                Select Case True
    '                    Case mintVFamilia <= 0 And mintVLinea <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " or (A.CodFamilia <> 0) and A.CodSubLinea is NULL)"
    '                    Case mintVFamilia > 0 And mintVLinea <= 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL))"
    '                    Case mintVFamilia > 0 And mintVLinea > 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL))"
    '                End Select
    '            Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS <= 0
    '                Joyeria
    '                cWHERE = cWHERE & " A.CodGrupo = " & gCODJOYERIA
    '                Select Case True
    '                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " and A.CodFamilia <> " & 0 & " and A.CodSubLinea is NOT NULL "
    '                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintJFamilia & " and A.CodSubLinea is NOT NULL "
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea is NOT NULL"
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea
    '                End Select
    '            Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS > 0
    '                Relojeria-Varios
    '                cWHERE = cWHERE & " A.CodGrupo <> " & gCODJOYERIA
    '                Select Case True
    '                    Case mintRMarca <= 0 And mintRModelo <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " and ((A.CodMarca <> " & 0 & ")"
    '                    Case mintRMarca > 0 And mintRModelo <= 0
    '                        cWHERE = cWHERE & " and ((A.CodMarca = " & mintRMarca & ")"
    '                    Case mintRMarca > 0 And mintRModelo > 0
    '                        cWHERE = cWHERE & " and ((A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & ")"
    '                End Select
    '                Select Case True
    '                    Case mintVFamilia <= 0 And mintVLinea <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " or (A.CodFamilia <> 0) and A.CodSubLinea is NULL)"
    '                    Case mintVFamilia > 0 And mintVLinea <= 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL))"
    '                    Case mintVFamilia > 0 And mintVLinea > 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL))"
    '                End Select
    '            Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS <= 0
    '                Relojeria
    '                cWHERE = cWHERE & " A.CodGrupo = " & gCODRELOJERIA
    '                Select Case True
    '                    Case mintRMarca <= 0 And mintRModelo <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " and A.CodMarca <> " & 0
    '                    Case mintRMarca > 0 And mintRModelo <= 0
    '                        cWHERE = cWHERE & " and A.CodMarca = " & mintRMarca
    '                    Case mintRMarca > 0 And mintRModelo > 0
    '                        cWHERE = cWHERE & " and A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo
    '                End Select
    '            Case nJOYERIA <= 0 And nRELOJERIA <= 0 And nVARIOS > 0
    '                Varios
    '                cWHERE = cWHERE & " A.CodGrupo = " & gCODVARIOS
    '                Select Case True
    '                    Case mintVFamilia <= 0 And mintVLinea <= 0
    '                        Todos
    '                        cWHERE = cWHERE & " and A.CodFamilia <> 0 "
    '                        'cWHERE = cWHERE & " and A.CodFamilia <> 0 and A.CodSubLinea is NULL "
    '                    Case mintVFamilia > 0 And mintVLinea <= 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintVFamilia
    '                        'cWHERE = cWHERE & " and A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL "
    '                    Case mintVFamilia > 0 And mintVLinea > 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea
    '                        'cWHERE = cWHERE & " and A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL "
    '                End Select
    '        End Select

    '        If chkRangoSinExistencia.Value = vbUnchecked And chkRangoTodos.Value = vbChecked Then
    '            cWHERE = cWHERE & " And (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados)) <> 0 "
    '        End If
    '        If chkRangoTodos.Value = vbUnchecked Then
    '            cWHERE = cWHERE & " And (SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados))  BetWeen " & Numerico(txtRangoDesde) & " and " & Numerico(txtRangoHasta)
    '        End If

    '        If dbcOrigen1 <> "" Then
    '            cWHERE = cWHERE & " And I.CodAlmacenOrigen = " & Numerico(txtCodOrigen)
    '        End If

    '        Agregar el Almacen al WHERE
    '        cWHERE = cWHERE & " And I.CodAlmacen= " & Numerico(txtCodSucursal) & "  "

    '        Select Case True
    '            Case optOrdArticulo.Value = True And cmbOrdArticulo.text = "Codigo"
    '                cORDERBY = " Order By I.CodArticulo "
    '            Case optOrdArticulo.Value = True And cmbOrdArticulo.text = "Descripcion"
    '                cORDERBY = " Order By A.DescArticulo "
    '            Case optOrdExistencia.Value = True And cmbOrdExistencia.text = "Ascendente"
    '                cORDERBY = " Order By Existencia Asc"
    '            Case optOrdExistencia.Value = True And cmbOrdExistencia.text = "Descendente"
    '                cORDERBY = " Order By Existencia Desc"
    '        End Select

    '        Armar el String  para el Select
    '        If mintJFamilia > 0 Or mintVFamilia > 0 Then
    '            cSELECT = cSELECT & " ISNULL(A.CodFamilia, 0) AS CodFamilia, ISNULL(Fa.DescFamilia, '') AS DescFamilia,  "
    '            cGROUPBY = cGROUPBY & " CodFamilia, "
    '        End If
    '        If mintJLinea > 0 Or mintVLinea > 0 Then
    '            cSELECT = cSELECT & " ISNULL(A.CodLinea, 0) AS CodLinea, ISNULL(Li.DescLinea, '') AS DescLinea,  "
    '            cGROUPBY = cGROUPBY & " CodLinea, "
    '        End If
    '        If mintJSubLinea > 0 Then
    '            cSELECT = cSELECT & " ISNULL(A.CodSubLinea,0) AS CodSublinea, ISNULL(Su.DescSubLinea, '') AS DescSubLinea,  "
    '            cGROUPBY = cGROUPBY & " CodSubLinea, "
    '        End If
    '        If mintRMarca > 0 Then
    '            cSELECT = cSELECT & " ISNULL(A.CodMarca, 0) AS codmarca, ISNULL(Ma.DescMarca, '') AS DescMarca,  "
    '            cGROUPBY = cGROUPBY & " CodMarca, "
    '        End If
    '        If mintRModelo > 0 Then
    '            cSELECT = cSELECT & " ISNULL(A.CodModelo, 0) AS codmodelo, ISNULL(Mo.DescModelo, '') AS DescModelo,  "
    '            cGROUPBY = cGROUPBY & " CodModelo, "
    '        End If

    '        DevuelveQuery = "FROM         dbo.Inventario I INNER JOIN " &
    '                    "dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " &
    '                    "dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " &
    '                    "dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen LEFT OUTER JOIN " &
    '                    "dbo.CatGrupos G ON A.CodGrupo = G.CodGrupo LEFT OUTER JOIN " &
    '                    "dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia LEFT OUTER JOIN " &
    '                    "dbo.CatLineas Li ON A.CodGrupo = Li.CodGrupo AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea LEFT OUTER JOIN " &
    '                    "dbo.CatSubLineas Su ON A.CodGrupo = Su.CodGrupo AND A.CodFamilia = Su.CodFamilia AND A.CodLinea = Su.CodLinea AND " &
    '                    "A.CodSubLinea = Su.CodSubLinea LEFT OUTER JOIN " &
    '                    "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca LEFT OUTER JOIN " &
    '                    "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND Mo.CodMarca = A.CodMarca AND A.CodModelo = Mo.CodModelo CROSS JOIN " &
    '                    "dbo.ConfiguracionGeneral " &
    '                "WHERE     (I.FechaMovto <= '" & Format(dtpFechaCorte, C_FORMATFECHAGUARDAR) & "') " &
    '                "GROUP BY I.CodArticulo, A.CodArticulo, A.DescArticulo, A.CodGrupo, A.CodUnidad, " &
    '                    "O.DescAlmacenOrigen, Al.DescAlmacen, I.CodAlmacen, I.CodAlmacenOrigen, A.CostoReal, " &
    '                    "dbo.ConfiguracionGeneral.NombreEmp, Al.DescAlmacen, O.DescAlmacenOrigen, A.DescArticulo, " &
    '                    "G.DescGrupo ,A.CodFamilia,A.CodSubLinea,A.CodMarca,A.CodSubLinea, Fa.DescFamilia, " &
    '                    "A.CodLinea,A.CodModelo,A.CodLinea , Li.DescLinea, Su.DescSubLinea, Ma.DescMarca,Mo.DescModelo, A.PrecioPubDolar,dbo.ConfiguracionGeneral.TasaIva,  " &
    '                    "A.CostoFacturaPesos , A.CostoAdicionalPesos ,A.CostoIndirectoPesos,PesosFijos " &
    '                cWHERE & " " &
    '                cORDERBY & "  "

    '        If optAlCosto.Value = True Then
    '            cSELECT = " SELECT  I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                        "AS Existencia, SUM(I.Apartados) AS Apartados, Round(A.CostoReal,2) AS CostoUnitario, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                        "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '        ElseIf optPrecioPublico.Value = True Then
    '            If chkIncluirIVA.Value = vbChecked Then
    '                cSELECT = " SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                                "AS Existencia, SUM(I.Apartados) AS Apartados, Case PesosFijos When 0 then Round(A.PrecioPubDolar,2) When 1 Then Round(A.PrecioPubDolar/" & gcurCorpoTIPOCAMBIODOLAR & ",2) End AS CostoUnitario,  " &
    '                                "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                                "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '            Else
    '                cSELECT = " SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                                "AS Existencia, SUM(I.Apartados) AS Apartados, Case PesosFijos When 0 then Round(A.PrecioPubDolar / (1 + dbo.ConfiguracionGeneral.TasaIva / 100) ,2) When 1 Then Round((A.PrecioPubDolar  / (1 + dbo.ConfiguracionGeneral.TasaIva / 100) )/" & gcurCorpoTIPOCAMBIODOLAR & ",2)  End AS CostoUnitario, " &
    '                                "LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                                "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '            End If
    '        Else
    '            Select Case True
    '                Case chkCostoFactura.Value = 1 And chkCostoAdicional.Value = 1 And chkCostoIndirecto.Value = 1
    '                    cSELECT = " SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                        "AS Existencia, SUM(I.Apartados) AS Apartados, Round(A.CostoFacturaPesos+ A.CostoAdicionalPesos + A.CostoIndirectoPesos,2) AS CostoUnitario, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                        "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '                Case chkCostoFactura.Value = 1 And chkCostoAdicional.Value = 1 And chkCostoIndirecto.Value = 0
    '                    cSELECT = " SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                        "AS Existencia, SUM(I.Apartados) AS Apartados, Round(A.CostoFacturaPesos+ A.CostoAdicionalPesos ,2) AS CostoUnitario, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                        "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '                Case chkCostoFactura.Value = 1 And chkCostoAdicional.Value = 0 And chkCostoIndirecto.Value = 1
    '                    cSELECT = " SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                        "AS Existencia, SUM(I.Apartados) AS Apartados, Round(A.CostoFacturaPesos+  A.CostoIndirectoPesos,2) AS CostoUnitario, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                        "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '                Case chkCostoFactura.Value = 1 And chkCostoAdicional.Value = 0 And chkCostoIndirecto.Value = 0
    '                    cSELECT = " SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                        "AS Existencia, SUM(I.Apartados) AS Apartados, Round(A.CostoFacturaPesos,2) AS CostoUnitario, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                        "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '                Case chkCostoFactura.Value = 0 And chkCostoAdicional.Value = 1 And chkCostoIndirecto.Value = 0
    '                    cSELECT = " SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                        "AS Existencia, SUM(I.Apartados) AS Apartados, Round(A.CostoAdicionalPesos,2) AS CostoUnitario, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                        "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '                Case chkCostoFactura.Value = 0 And chkCostoAdicional.Value = 0 And chkCostoIndirecto.Value = 1
    '                    cSELECT = " SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                        "AS Existencia, SUM(I.Apartados) AS Apartados, Round(A.CostoIndirectoPesos,2) AS CostoUnitario, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                        "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '                Case chkCostoFactura.Value = 0 And chkCostoAdicional.Value = 1 And chkCostoIndirecto.Value = 1
    '                    cSELECT = " SELECT     I.CodAlmacen, I.CodAlmacenOrigen, A.CodGrupo, A.CodArticulo, SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) " &
    '                        "AS Existencia, SUM(I.Apartados) AS Apartados, Round(A.CostoAdicionalPesos + A.CostoIndirectoPesos,2) AS CostoUnitario, LTRIM(RTRIM(dbo.ConfiguracionGeneral.NombreEmp)) AS NombreEmp, " &
    '                        "Al.DescAlmacen , O.DescAlmacenOrigen, A.DescArticulo, " & cSELECT & "  G.DescGrupo  "
    '                Case chkCostoFactura.Value = 0 And chkCostoAdicional.Value = 0 And chkCostoIndirecto.Value = 0
    '                    MsgBox "Debe elegir, por lo menos, un costo en pesos para poder generar el reporte", vbOKOnly + vbInformation, gstrCorpoNOMBREEMPRESA
    '                     DevuelveQuery = ""
    '                    Exit Function
    '            End Select

    '        End If
    '        DevuelveQuery = cSELECT & DevuelveQuery

    '        Exit Function

    'MErr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '    End Function


    Private Sub txtRangoDesde_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRangoDesde.Enter
        SelTextoTxt(txtRangoDesde)
    End Sub

    Private Sub txtRangoDesde_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRangoDesde.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub txtRangoHasta_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRangoHasta.Enter
        SelTextoTxt(txtRangoHasta)
    End Sub


    Sub PonerCodigoOrigen()
        gStrSql = "SELECT CodAlmacenOrigen,LTRIM(RTRIM(DescAlmacenOrigen)) as DescAlmacen FROM CatOrigen WHERE DescAlmacenOrigen LIKE '" & Trim(dbcOrigen1.Text) & "'  ORDER BY DescAlmacenOrigen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'DCLostFocus dbcAlmacenSalida, gStrSql, intCodSucursal
        mblnFueraChange = True
        If RsGral.RecordCount <= 0 Then
            txtCodOrigen.Text = ""
        Else
            txtCodOrigen.Text = RsGral.Fields("CodAlmacenOrigen").Value
        End If
        mblnFueraChange = False
    End Sub

    Sub PonerCodigoSucursal()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_SELECT_DATOS"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = Cmd.Execute

        'DCLostFocus dbcAlmacenSalida, gStrSql, intCodSucursal
        mblnFueraChange = True
        If RsGral.RecordCount <= 0 Then
            txtCodSucursal.Text = ""
        Else
            txtCodSucursal.Text = VB6.Format(RsGral.Fields("CodAlmacen").Value, "000")
        End If
        mblnFueraChange = False
    End Sub


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtRangoDesde = New System.Windows.Forms.TextBox()
        Me.txtRangoHasta = New System.Windows.Forms.TextBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Frame5 = New System.Windows.Forms.GroupBox()
        Me.optTotal = New System.Windows.Forms.RadioButton()
        Me.optxSuc = New System.Windows.Forms.RadioButton()
        Me.fraRangoExistencia = New System.Windows.Forms.GroupBox()
        Me.chkRangoSinExistencia = New System.Windows.Forms.CheckBox()
        Me.chkRangoTodos = New System.Windows.Forms.CheckBox()
        Me.lblHasta = New System.Windows.Forms.Label()
        Me.lblDesde = New System.Windows.Forms.Label()
        Me.fraOrdenamiento = New System.Windows.Forms.GroupBox()
        Me.cmbOrdArticulo = New System.Windows.Forms.ComboBox()
        Me.cmbOrdExistencia = New System.Windows.Forms.ComboBox()
        Me.optOrdExistencia = New System.Windows.Forms.RadioButton()
        Me.optOrdArticulo = New System.Windows.Forms.RadioButton()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.chkCostoFactura = New System.Windows.Forms.CheckBox()
        Me.chkCostoIndirecto = New System.Windows.Forms.CheckBox()
        Me.chkCostoAdicional = New System.Windows.Forms.CheckBox()
        Me.chkIncluirIVA = New System.Windows.Forms.CheckBox()
        Me.optAlCosto = New System.Windows.Forms.RadioButton()
        Me.optPrecioPublico = New System.Windows.Forms.RadioButton()
        Me.optUltimoCostoPesos = New System.Windows.Forms.RadioButton()
        Me.chkMostrarAparatdos = New System.Windows.Forms.CheckBox()
        Me.fraGrupo = New System.Windows.Forms.GroupBox()
        Me.chkRelojeria = New System.Windows.Forms.CheckBox()
        Me.chkVarios = New System.Windows.Forms.CheckBox()
        Me.chkJoyeria = New System.Windows.Forms.CheckBox()
        Me._Frame3_0 = New System.Windows.Forms.GroupBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.dbcJFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcJLinea = New System.Windows.Forms.ComboBox()
        Me.dbcJSubLinea = New System.Windows.Forms.ComboBox()
        Me.dbcVLinea = New System.Windows.Forms.ComboBox()
        Me.dbcRMarca = New System.Windows.Forms.ComboBox()
        Me.dbcRModelo = New System.Windows.Forms.ComboBox()
        Me.dbcVFamilia = New System.Windows.Forms.ComboBox()
        Me._lblVentas_8 = New System.Windows.Forms.Label()
        Me._lblVentas_7 = New System.Windows.Forms.Label()
        Me._lblVentas_6 = New System.Windows.Forms.Label()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me._lblVentas_4 = New System.Windows.Forms.Label()
        Me._lblVentas_3 = New System.Windows.Forms.Label()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me.txtCodOrigen = New System.Windows.Forms.TextBox()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me.dbcOrigen1 = New System.Windows.Forms.ComboBox()
        Me.dtpFechaCorte = New System.Windows.Forms.DateTimePicker()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.Frame3 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.fraRangoExistencia.SuspendLayout()
        Me.fraOrdenamiento.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.fraGrupo.SuspendLayout()
        CType(Me.Frame3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtRangoDesde
        '
        Me.txtRangoDesde.AcceptsReturn = True
        Me.txtRangoDesde.BackColor = System.Drawing.SystemColors.Window
        Me.txtRangoDesde.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRangoDesde.Enabled = False
        Me.txtRangoDesde.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRangoDesde.Location = New System.Drawing.Point(192, 24)
        Me.txtRangoDesde.MaxLength = 8
        Me.txtRangoDesde.Name = "txtRangoDesde"
        Me.txtRangoDesde.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRangoDesde.Size = New System.Drawing.Size(65, 21)
        Me.txtRangoDesde.TabIndex = 36
        Me.txtRangoDesde.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtRangoDesde, "Rango Inferior a Mostrar")
        '
        'txtRangoHasta
        '
        Me.txtRangoHasta.AcceptsReturn = True
        Me.txtRangoHasta.BackColor = System.Drawing.SystemColors.Window
        Me.txtRangoHasta.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRangoHasta.Enabled = False
        Me.txtRangoHasta.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRangoHasta.Location = New System.Drawing.Point(192, 48)
        Me.txtRangoHasta.MaxLength = 8
        Me.txtRangoHasta.Name = "txtRangoHasta"
        Me.txtRangoHasta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRangoHasta.Size = New System.Drawing.Size(65, 21)
        Me.txtRangoHasta.TabIndex = 38
        Me.txtRangoHasta.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.txtRangoHasta, "Rango Superior a Mostrar")
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Frame5)
        Me.Frame1.Controls.Add(Me.optTotal)
        Me.Frame1.Controls.Add(Me.optxSuc)
        Me.Frame1.Controls.Add(Me.fraRangoExistencia)
        Me.Frame1.Controls.Add(Me.fraOrdenamiento)
        Me.Frame1.Controls.Add(Me.Frame2)
        Me.Frame1.Controls.Add(Me.chkMostrarAparatdos)
        Me.Frame1.Controls.Add(Me.fraGrupo)
        Me.Frame1.Controls.Add(Me.txtCodOrigen)
        Me.Frame1.Controls.Add(Me.txtCodSucursal)
        Me.Frame1.Controls.Add(Me.dbcSucursales)
        Me.Frame1.Controls.Add(Me.dbcOrigen1)
        Me.Frame1.Controls.Add(Me.dtpFechaCorte)
        Me.Frame1.Controls.Add(Me._Label1_1)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me._Label1_0)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(449, 616)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'Frame5
        '
        Me.Frame5.BackColor = System.Drawing.SystemColors.Control
        Me.Frame5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame5.Location = New System.Drawing.Point(12, 67)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame5.Size = New System.Drawing.Size(425, 4)
        Me.Frame5.TabIndex = 52
        Me.Frame5.TabStop = False
        '
        'optTotal
        '
        Me.optTotal.BackColor = System.Drawing.SystemColors.Control
        Me.optTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.optTotal.ForeColor = System.Drawing.Color.Black
        Me.optTotal.Location = New System.Drawing.Point(329, 50)
        Me.optTotal.Name = "optTotal"
        Me.optTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optTotal.Size = New System.Drawing.Size(76, 21)
        Me.optTotal.TabIndex = 4
        Me.optTotal.TabStop = True
        Me.optTotal.Text = "Total"
        Me.optTotal.UseVisualStyleBackColor = False
        '
        'optxSuc
        '
        Me.optxSuc.BackColor = System.Drawing.SystemColors.Control
        Me.optxSuc.Cursor = System.Windows.Forms.Cursors.Default
        Me.optxSuc.ForeColor = System.Drawing.Color.Black
        Me.optxSuc.Location = New System.Drawing.Point(30, 50)
        Me.optxSuc.Name = "optxSuc"
        Me.optxSuc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optxSuc.Size = New System.Drawing.Size(124, 21)
        Me.optxSuc.TabIndex = 3
        Me.optxSuc.TabStop = True
        Me.optxSuc.Text = "Por Sucursal"
        Me.optxSuc.UseVisualStyleBackColor = False
        '
        'fraRangoExistencia
        '
        Me.fraRangoExistencia.BackColor = System.Drawing.SystemColors.Control
        Me.fraRangoExistencia.Controls.Add(Me.chkRangoSinExistencia)
        Me.fraRangoExistencia.Controls.Add(Me.chkRangoTodos)
        Me.fraRangoExistencia.Controls.Add(Me.txtRangoDesde)
        Me.fraRangoExistencia.Controls.Add(Me.txtRangoHasta)
        Me.fraRangoExistencia.Controls.Add(Me.lblHasta)
        Me.fraRangoExistencia.Controls.Add(Me.lblDesde)
        Me.fraRangoExistencia.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraRangoExistencia.Location = New System.Drawing.Point(12, 410)
        Me.fraRangoExistencia.Name = "fraRangoExistencia"
        Me.fraRangoExistencia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRangoExistencia.Size = New System.Drawing.Size(265, 77)
        Me.fraRangoExistencia.TabIndex = 32
        Me.fraRangoExistencia.TabStop = False
        Me.fraRangoExistencia.Text = " Rango de Existencias "
        '
        'chkRangoSinExistencia
        '
        Me.chkRangoSinExistencia.BackColor = System.Drawing.SystemColors.Control
        Me.chkRangoSinExistencia.Checked = True
        Me.chkRangoSinExistencia.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRangoSinExistencia.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkRangoSinExistencia.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkRangoSinExistencia.Location = New System.Drawing.Point(8, 48)
        Me.chkRangoSinExistencia.Name = "chkRangoSinExistencia"
        Me.chkRangoSinExistencia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkRangoSinExistencia.Size = New System.Drawing.Size(121, 17)
        Me.chkRangoSinExistencia.TabIndex = 34
        Me.chkRangoSinExistencia.Text = "Incluir sin Existencia"
        Me.chkRangoSinExistencia.UseVisualStyleBackColor = False
        '
        'chkRangoTodos
        '
        Me.chkRangoTodos.BackColor = System.Drawing.SystemColors.Control
        Me.chkRangoTodos.Checked = True
        Me.chkRangoTodos.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRangoTodos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkRangoTodos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkRangoTodos.Location = New System.Drawing.Point(8, 24)
        Me.chkRangoTodos.Name = "chkRangoTodos"
        Me.chkRangoTodos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkRangoTodos.Size = New System.Drawing.Size(121, 17)
        Me.chkRangoTodos.TabIndex = 33
        Me.chkRangoTodos.Text = "Todos los rangos"
        Me.chkRangoTodos.UseVisualStyleBackColor = False
        '
        'lblHasta
        '
        Me.lblHasta.AutoSize = True
        Me.lblHasta.BackColor = System.Drawing.SystemColors.Control
        Me.lblHasta.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblHasta.Enabled = False
        Me.lblHasta.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblHasta.Location = New System.Drawing.Point(144, 48)
        Me.lblHasta.Name = "lblHasta"
        Me.lblHasta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblHasta.Size = New System.Drawing.Size(35, 13)
        Me.lblHasta.TabIndex = 37
        Me.lblHasta.Text = "Hasta"
        '
        'lblDesde
        '
        Me.lblDesde.AutoSize = True
        Me.lblDesde.BackColor = System.Drawing.SystemColors.Control
        Me.lblDesde.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDesde.Enabled = False
        Me.lblDesde.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDesde.Location = New System.Drawing.Point(144, 24)
        Me.lblDesde.Name = "lblDesde"
        Me.lblDesde.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDesde.Size = New System.Drawing.Size(38, 13)
        Me.lblDesde.TabIndex = 35
        Me.lblDesde.Text = "Desde"
        '
        'fraOrdenamiento
        '
        Me.fraOrdenamiento.BackColor = System.Drawing.SystemColors.Control
        Me.fraOrdenamiento.Controls.Add(Me.cmbOrdArticulo)
        Me.fraOrdenamiento.Controls.Add(Me.cmbOrdExistencia)
        Me.fraOrdenamiento.Controls.Add(Me.optOrdExistencia)
        Me.fraOrdenamiento.Controls.Add(Me.optOrdArticulo)
        Me.fraOrdenamiento.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.fraOrdenamiento.Location = New System.Drawing.Point(12, 495)
        Me.fraOrdenamiento.Name = "fraOrdenamiento"
        Me.fraOrdenamiento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraOrdenamiento.Size = New System.Drawing.Size(265, 89)
        Me.fraOrdenamiento.TabIndex = 39
        Me.fraOrdenamiento.TabStop = False
        Me.fraOrdenamiento.Text = "Ordenamiento "
        '
        'cmbOrdArticulo
        '
        Me.cmbOrdArticulo.BackColor = System.Drawing.SystemColors.Window
        Me.cmbOrdArticulo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbOrdArticulo.Enabled = False
        Me.cmbOrdArticulo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbOrdArticulo.Items.AddRange(New Object() {"Codigo", "Descripcion"})
        Me.cmbOrdArticulo.Location = New System.Drawing.Point(104, 24)
        Me.cmbOrdArticulo.Name = "cmbOrdArticulo"
        Me.cmbOrdArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbOrdArticulo.Size = New System.Drawing.Size(145, 21)
        Me.cmbOrdArticulo.TabIndex = 41
        '
        'cmbOrdExistencia
        '
        Me.cmbOrdExistencia.BackColor = System.Drawing.SystemColors.Window
        Me.cmbOrdExistencia.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbOrdExistencia.Enabled = False
        Me.cmbOrdExistencia.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbOrdExistencia.Items.AddRange(New Object() {"Ascendente", "Descendente"})
        Me.cmbOrdExistencia.Location = New System.Drawing.Point(104, 56)
        Me.cmbOrdExistencia.Name = "cmbOrdExistencia"
        Me.cmbOrdExistencia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbOrdExistencia.Size = New System.Drawing.Size(145, 21)
        Me.cmbOrdExistencia.TabIndex = 43
        '
        'optOrdExistencia
        '
        Me.optOrdExistencia.BackColor = System.Drawing.SystemColors.Control
        Me.optOrdExistencia.Cursor = System.Windows.Forms.Cursors.Default
        Me.optOrdExistencia.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optOrdExistencia.Location = New System.Drawing.Point(8, 56)
        Me.optOrdExistencia.Name = "optOrdExistencia"
        Me.optOrdExistencia.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optOrdExistencia.Size = New System.Drawing.Size(93, 17)
        Me.optOrdExistencia.TabIndex = 42
        Me.optOrdExistencia.TabStop = True
        Me.optOrdExistencia.Text = "Por Existencia"
        Me.optOrdExistencia.UseVisualStyleBackColor = False
        '
        'optOrdArticulo
        '
        Me.optOrdArticulo.BackColor = System.Drawing.SystemColors.Control
        Me.optOrdArticulo.Cursor = System.Windows.Forms.Cursors.Default
        Me.optOrdArticulo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optOrdArticulo.Location = New System.Drawing.Point(8, 24)
        Me.optOrdArticulo.Name = "optOrdArticulo"
        Me.optOrdArticulo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optOrdArticulo.Size = New System.Drawing.Size(93, 17)
        Me.optOrdArticulo.TabIndex = 40
        Me.optOrdArticulo.TabStop = True
        Me.optOrdArticulo.Text = "Por Artículo"
        Me.optOrdArticulo.UseVisualStyleBackColor = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.chkCostoFactura)
        Me.Frame2.Controls.Add(Me.chkCostoIndirecto)
        Me.Frame2.Controls.Add(Me.chkCostoAdicional)
        Me.Frame2.Controls.Add(Me.chkIncluirIVA)
        Me.Frame2.Controls.Add(Me.optAlCosto)
        Me.Frame2.Controls.Add(Me.optPrecioPublico)
        Me.Frame2.Controls.Add(Me.optUltimoCostoPesos)
        Me.Frame2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Frame2.Location = New System.Drawing.Point(284, 410)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(153, 191)
        Me.Frame2.TabIndex = 44
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Imprimir Artículos ..."
        '
        'chkCostoFactura
        '
        Me.chkCostoFactura.BackColor = System.Drawing.SystemColors.Control
        Me.chkCostoFactura.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCostoFactura.Enabled = False
        Me.chkCostoFactura.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCostoFactura.Location = New System.Drawing.Point(48, 104)
        Me.chkCostoFactura.Name = "chkCostoFactura"
        Me.chkCostoFactura.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCostoFactura.Size = New System.Drawing.Size(97, 29)
        Me.chkCostoFactura.TabIndex = 49
        Me.chkCostoFactura.Text = "Costo Factura"
        Me.chkCostoFactura.UseVisualStyleBackColor = False
        '
        'chkCostoIndirecto
        '
        Me.chkCostoIndirecto.BackColor = System.Drawing.SystemColors.Control
        Me.chkCostoIndirecto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCostoIndirecto.Enabled = False
        Me.chkCostoIndirecto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCostoIndirecto.Location = New System.Drawing.Point(48, 147)
        Me.chkCostoIndirecto.Name = "chkCostoIndirecto"
        Me.chkCostoIndirecto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCostoIndirecto.Size = New System.Drawing.Size(97, 38)
        Me.chkCostoIndirecto.TabIndex = 51
        Me.chkCostoIndirecto.Text = "Costo Indirecto"
        Me.chkCostoIndirecto.UseVisualStyleBackColor = False
        '
        'chkCostoAdicional
        '
        Me.chkCostoAdicional.BackColor = System.Drawing.SystemColors.Control
        Me.chkCostoAdicional.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCostoAdicional.Enabled = False
        Me.chkCostoAdicional.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCostoAdicional.Location = New System.Drawing.Point(48, 124)
        Me.chkCostoAdicional.Name = "chkCostoAdicional"
        Me.chkCostoAdicional.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCostoAdicional.Size = New System.Drawing.Size(103, 34)
        Me.chkCostoAdicional.TabIndex = 50
        Me.chkCostoAdicional.Text = "Costo Adicional"
        Me.chkCostoAdicional.UseVisualStyleBackColor = False
        '
        'chkIncluirIVA
        '
        Me.chkIncluirIVA.BackColor = System.Drawing.SystemColors.Control
        Me.chkIncluirIVA.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncluirIVA.Enabled = False
        Me.chkIncluirIVA.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIncluirIVA.Location = New System.Drawing.Point(45, 59)
        Me.chkIncluirIVA.Name = "chkIncluirIVA"
        Me.chkIncluirIVA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncluirIVA.Size = New System.Drawing.Size(95, 21)
        Me.chkIncluirIVA.TabIndex = 47
        Me.chkIncluirIVA.Text = "Incluir IVA"
        Me.chkIncluirIVA.UseVisualStyleBackColor = False
        '
        'optAlCosto
        '
        Me.optAlCosto.BackColor = System.Drawing.SystemColors.Control
        Me.optAlCosto.Checked = True
        Me.optAlCosto.Cursor = System.Windows.Forms.Cursors.Default
        Me.optAlCosto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optAlCosto.Location = New System.Drawing.Point(16, 16)
        Me.optAlCosto.Name = "optAlCosto"
        Me.optAlCosto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optAlCosto.Size = New System.Drawing.Size(105, 21)
        Me.optAlCosto.TabIndex = 45
        Me.optAlCosto.TabStop = True
        Me.optAlCosto.Text = "Al Costo"
        Me.optAlCosto.UseVisualStyleBackColor = False
        '
        'optPrecioPublico
        '
        Me.optPrecioPublico.BackColor = System.Drawing.SystemColors.Control
        Me.optPrecioPublico.Cursor = System.Windows.Forms.Cursors.Default
        Me.optPrecioPublico.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optPrecioPublico.Location = New System.Drawing.Point(16, 37)
        Me.optPrecioPublico.Name = "optPrecioPublico"
        Me.optPrecioPublico.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optPrecioPublico.Size = New System.Drawing.Size(113, 29)
        Me.optPrecioPublico.TabIndex = 46
        Me.optPrecioPublico.TabStop = True
        Me.optPrecioPublico.Text = "Precio Público"
        Me.optPrecioPublico.UseVisualStyleBackColor = False
        '
        'optUltimoCostoPesos
        '
        Me.optUltimoCostoPesos.BackColor = System.Drawing.SystemColors.Control
        Me.optUltimoCostoPesos.Cursor = System.Windows.Forms.Cursors.Default
        Me.optUltimoCostoPesos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUltimoCostoPesos.Location = New System.Drawing.Point(16, 83)
        Me.optUltimoCostoPesos.Name = "optUltimoCostoPesos"
        Me.optUltimoCostoPesos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optUltimoCostoPesos.Size = New System.Drawing.Size(127, 29)
        Me.optUltimoCostoPesos.TabIndex = 48
        Me.optUltimoCostoPesos.TabStop = True
        Me.optUltimoCostoPesos.Text = "Último costo pesos"
        Me.optUltimoCostoPesos.UseVisualStyleBackColor = False
        '
        'chkMostrarAparatdos
        '
        Me.chkMostrarAparatdos.BackColor = System.Drawing.SystemColors.Control
        Me.chkMostrarAparatdos.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarAparatdos.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMostrarAparatdos.Location = New System.Drawing.Point(12, 386)
        Me.chkMostrarAparatdos.Name = "chkMostrarAparatdos"
        Me.chkMostrarAparatdos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarAparatdos.Size = New System.Drawing.Size(129, 17)
        Me.chkMostrarAparatdos.TabIndex = 31
        Me.chkMostrarAparatdos.Text = "Mostrar Apartados"
        Me.chkMostrarAparatdos.UseVisualStyleBackColor = False
        '
        'fraGrupo
        '
        Me.fraGrupo.BackColor = System.Drawing.SystemColors.Control
        Me.fraGrupo.Controls.Add(Me.chkRelojeria)
        Me.fraGrupo.Controls.Add(Me.chkVarios)
        Me.fraGrupo.Controls.Add(Me.chkJoyeria)
        Me.fraGrupo.Controls.Add(Me._Frame3_0)
        Me.fraGrupo.Controls.Add(Me.Frame4)
        Me.fraGrupo.Controls.Add(Me.dbcJFamilia)
        Me.fraGrupo.Controls.Add(Me.dbcJLinea)
        Me.fraGrupo.Controls.Add(Me.dbcJSubLinea)
        Me.fraGrupo.Controls.Add(Me.dbcVLinea)
        Me.fraGrupo.Controls.Add(Me.dbcRMarca)
        Me.fraGrupo.Controls.Add(Me.dbcRModelo)
        Me.fraGrupo.Controls.Add(Me.dbcVFamilia)
        Me.fraGrupo.Controls.Add(Me._lblVentas_8)
        Me.fraGrupo.Controls.Add(Me._lblVentas_7)
        Me.fraGrupo.Controls.Add(Me._lblVentas_6)
        Me.fraGrupo.Controls.Add(Me._lblVentas_5)
        Me.fraGrupo.Controls.Add(Me._lblVentas_4)
        Me.fraGrupo.Controls.Add(Me._lblVentas_3)
        Me.fraGrupo.Controls.Add(Me._lblVentas_0)
        Me.fraGrupo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGrupo.Location = New System.Drawing.Point(12, 143)
        Me.fraGrupo.Name = "fraGrupo"
        Me.fraGrupo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGrupo.Size = New System.Drawing.Size(425, 233)
        Me.fraGrupo.TabIndex = 11
        Me.fraGrupo.TabStop = False
        Me.fraGrupo.Text = " Grupo "
        '
        'chkRelojeria
        '
        Me.chkRelojeria.BackColor = System.Drawing.SystemColors.Control
        Me.chkRelojeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkRelojeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkRelojeria.Location = New System.Drawing.Point(20, 112)
        Me.chkRelojeria.Name = "chkRelojeria"
        Me.chkRelojeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkRelojeria.Size = New System.Drawing.Size(81, 17)
        Me.chkRelojeria.TabIndex = 20
        Me.chkRelojeria.Text = "Relojería"
        Me.chkRelojeria.UseVisualStyleBackColor = False
        '
        'chkVarios
        '
        Me.chkVarios.BackColor = System.Drawing.SystemColors.Control
        Me.chkVarios.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVarios.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkVarios.Location = New System.Drawing.Point(20, 176)
        Me.chkVarios.Name = "chkVarios"
        Me.chkVarios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVarios.Size = New System.Drawing.Size(81, 17)
        Me.chkVarios.TabIndex = 26
        Me.chkVarios.Text = "Varios"
        Me.chkVarios.UseVisualStyleBackColor = False
        '
        'chkJoyeria
        '
        Me.chkJoyeria.BackColor = System.Drawing.SystemColors.Control
        Me.chkJoyeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkJoyeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkJoyeria.Location = New System.Drawing.Point(20, 24)
        Me.chkJoyeria.Name = "chkJoyeria"
        Me.chkJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkJoyeria.Size = New System.Drawing.Size(81, 17)
        Me.chkJoyeria.TabIndex = 12
        Me.chkJoyeria.Text = "Joyería"
        Me.chkJoyeria.UseVisualStyleBackColor = False
        '
        '_Frame3_0
        '
        Me._Frame3_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame3_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.SetIndex(Me._Frame3_0, CType(0, Short))
        Me._Frame3_0.Location = New System.Drawing.Point(16, 96)
        Me._Frame3_0.Name = "_Frame3_0"
        Me._Frame3_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame3_0.Size = New System.Drawing.Size(401, 2)
        Me._Frame3_0.TabIndex = 19
        Me._Frame3_0.TabStop = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(16, 160)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(401, 2)
        Me.Frame4.TabIndex = 25
        Me.Frame4.TabStop = False
        '
        'dbcJFamilia
        '
        Me.dbcJFamilia.Location = New System.Drawing.Point(162, 24)
        Me.dbcJFamilia.Name = "dbcJFamilia"
        Me.dbcJFamilia.Size = New System.Drawing.Size(253, 21)
        Me.dbcJFamilia.TabIndex = 14
        '
        'dbcJLinea
        '
        Me.dbcJLinea.Location = New System.Drawing.Point(162, 48)
        Me.dbcJLinea.Name = "dbcJLinea"
        Me.dbcJLinea.Size = New System.Drawing.Size(253, 21)
        Me.dbcJLinea.TabIndex = 16
        '
        'dbcJSubLinea
        '
        Me.dbcJSubLinea.Location = New System.Drawing.Point(162, 72)
        Me.dbcJSubLinea.Name = "dbcJSubLinea"
        Me.dbcJSubLinea.Size = New System.Drawing.Size(253, 21)
        Me.dbcJSubLinea.TabIndex = 18
        '
        'dbcVLinea
        '
        Me.dbcVLinea.Location = New System.Drawing.Point(162, 200)
        Me.dbcVLinea.Name = "dbcVLinea"
        Me.dbcVLinea.Size = New System.Drawing.Size(253, 21)
        Me.dbcVLinea.TabIndex = 30
        '
        'dbcRMarca
        '
        Me.dbcRMarca.Location = New System.Drawing.Point(162, 112)
        Me.dbcRMarca.Name = "dbcRMarca"
        Me.dbcRMarca.Size = New System.Drawing.Size(253, 21)
        Me.dbcRMarca.TabIndex = 22
        '
        'dbcRModelo
        '
        Me.dbcRModelo.Location = New System.Drawing.Point(162, 136)
        Me.dbcRModelo.Name = "dbcRModelo"
        Me.dbcRModelo.Size = New System.Drawing.Size(253, 21)
        Me.dbcRModelo.TabIndex = 24
        '
        'dbcVFamilia
        '
        Me.dbcVFamilia.Location = New System.Drawing.Point(162, 176)
        Me.dbcVFamilia.Name = "dbcVFamilia"
        Me.dbcVFamilia.Size = New System.Drawing.Size(253, 21)
        Me.dbcVFamilia.TabIndex = 28
        '
        '_lblVentas_8
        '
        Me._lblVentas_8.AutoSize = True
        Me._lblVentas_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_8, CType(8, Short))
        Me._lblVentas_8.Location = New System.Drawing.Point(100, 199)
        Me._lblVentas_8.Name = "_lblVentas_8"
        Me._lblVentas_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_8.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_8.TabIndex = 29
        Me._lblVentas_8.Text = "Línea"
        '
        '_lblVentas_7
        '
        Me._lblVentas_7.AutoSize = True
        Me._lblVentas_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_7, CType(7, Short))
        Me._lblVentas_7.Location = New System.Drawing.Point(100, 176)
        Me._lblVentas_7.Name = "_lblVentas_7"
        Me._lblVentas_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_7.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_7.TabIndex = 27
        Me._lblVentas_7.Text = "Familia"
        '
        '_lblVentas_6
        '
        Me._lblVentas_6.AutoSize = True
        Me._lblVentas_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_6, CType(6, Short))
        Me._lblVentas_6.Location = New System.Drawing.Point(100, 136)
        Me._lblVentas_6.Name = "_lblVentas_6"
        Me._lblVentas_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_6.Size = New System.Drawing.Size(42, 13)
        Me._lblVentas_6.TabIndex = 23
        Me._lblVentas_6.Text = "Modelo"
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.AutoSize = True
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_5, CType(5, Short))
        Me._lblVentas_5.Location = New System.Drawing.Point(100, 112)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(37, 13)
        Me._lblVentas_5.TabIndex = 21
        Me._lblVentas_5.Text = "Marca"
        '
        '_lblVentas_4
        '
        Me._lblVentas_4.AutoSize = True
        Me._lblVentas_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_4, CType(4, Short))
        Me._lblVentas_4.Location = New System.Drawing.Point(100, 74)
        Me._lblVentas_4.Name = "_lblVentas_4"
        Me._lblVentas_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_4.Size = New System.Drawing.Size(54, 13)
        Me._lblVentas_4.TabIndex = 17
        Me._lblVentas_4.Text = "SubLínea"
        '
        '_lblVentas_3
        '
        Me._lblVentas_3.AutoSize = True
        Me._lblVentas_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_3, CType(3, Short))
        Me._lblVentas_3.Location = New System.Drawing.Point(100, 49)
        Me._lblVentas_3.Name = "_lblVentas_3"
        Me._lblVentas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_3.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_3.TabIndex = 15
        Me._lblVentas_3.Text = "Línea"
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVentas.SetIndex(Me._lblVentas_0, CType(0, Short))
        Me._lblVentas_0.Location = New System.Drawing.Point(100, 24)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_0.TabIndex = 13
        Me._lblVentas_0.Text = "Familia"
        '
        'txtCodOrigen
        '
        Me.txtCodOrigen.AcceptsReturn = True
        Me.txtCodOrigen.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodOrigen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodOrigen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodOrigen.Location = New System.Drawing.Point(102, 110)
        Me.txtCodOrigen.MaxLength = 10
        Me.txtCodOrigen.Name = "txtCodOrigen"
        Me.txtCodOrigen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodOrigen.Size = New System.Drawing.Size(49, 21)
        Me.txtCodOrigen.TabIndex = 9
        Me.txtCodOrigen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(102, 86)
        Me.txtCodSucursal.MaxLength = 10
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(49, 21)
        Me.txtCodSucursal.TabIndex = 6
        Me.txtCodSucursal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(160, 86)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(275, 21)
        Me.dbcSucursales.TabIndex = 7
        '
        'dbcOrigen1
        '
        Me.dbcOrigen1.Location = New System.Drawing.Point(160, 110)
        Me.dbcOrigen1.Name = "dbcOrigen1"
        Me.dbcOrigen1.Size = New System.Drawing.Size(275, 21)
        Me.dbcOrigen1.TabIndex = 10
        '
        'dtpFechaCorte
        '
        Me.dtpFechaCorte.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFechaCorte.Location = New System.Drawing.Point(328, 16)
        Me.dtpFechaCorte.Name = "dtpFechaCorte"
        Me.dtpFechaCorte.Size = New System.Drawing.Size(105, 20)
        Me.dtpFechaCorte.TabIndex = 2
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_1, CType(1, Short))
        Me._Label1_1.Location = New System.Drawing.Point(222, 22)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(105, 17)
        Me._Label1_1.TabIndex = 1
        Me._Label1_1.Text = "Fecha de Corte :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(30, 110)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Origen :"
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.SetIndex(Me._Label1_0, CType(0, Short))
        Me._Label1_0.Location = New System.Drawing.Point(30, 86)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(73, 17)
        Me._Label1_0.TabIndex = 5
        Me._Label1_0.Text = "Sucursal :"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(123, 639)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 94
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(8, 639)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 93
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'btnBuscar
        '
        Me.btnBuscar.Location = New System.Drawing.Point(238, 640)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(109, 36)
        Me.btnBuscar.TabIndex = 92
        Me.btnBuscar.Text = "&Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = False
        '
        'frmRptExistenciasyCostos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(465, 687)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(214, 69)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRptExistenciasyCostos"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Existencias y Costos"
        Me.Frame1.ResumeLayout(False)
        Me.fraRangoExistencia.ResumeLayout(False)
        Me.fraRangoExistencia.PerformLayout()
        Me.fraOrdenamiento.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.fraGrupo.ResumeLayout(False)
        Me.fraGrupo.PerformLayout()
        CType(Me.Frame3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

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