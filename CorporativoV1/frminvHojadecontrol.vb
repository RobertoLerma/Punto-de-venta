Option Strict Off
Option Explicit On
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Public Class frminvHojadecontrol
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    'Programa: Hoja de Control de Inventarios
    'Autor: Rosaura Torres López
    'Fecha de Creación: 26/08/03
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkExistenciaMayorCero As System.Windows.Forms.CheckBox
    Public WithEvents chkIncluirExistenciaTeorica As System.Windows.Forms.CheckBox
    Public WithEvents chkRelojeria As System.Windows.Forms.CheckBox
    Public WithEvents chkVarios As System.Windows.Forms.CheckBox
    Public WithEvents chkJoyeria As System.Windows.Forms.CheckBox
    Public WithEvents _Frame3_0 As System.Windows.Forms.GroupBox
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
    Public WithEvents txtCodSucursal As System.Windows.Forms.TextBox
    Public WithEvents txtCodOrigen As System.Windows.Forms.TextBox
    Public WithEvents _Frame3_1 As System.Windows.Forms.GroupBox
    Public WithEvents chkOrdenarporGrupo As System.Windows.Forms.CheckBox
    Public WithEvents optDescripcion As System.Windows.Forms.RadioButton
    Public WithEvents optCodigo As System.Windows.Forms.RadioButton
    Public WithEvents optCodActual As System.Windows.Forms.RadioButton
    Public WithEvents optCodAnterior As System.Windows.Forms.RadioButton
    Public WithEvents fraCodigo As System.Windows.Forms.Panel
    Public WithEvents fraOrdenamiento As System.Windows.Forms.GroupBox
    Public WithEvents dbcJFamilia As System.Windows.Forms.ComboBox
    Public WithEvents dbcJLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcJSubLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcVLinea As System.Windows.Forms.ComboBox
    Public WithEvents dbcRMarca As System.Windows.Forms.ComboBox
    Public WithEvents dbcRModelo As System.Windows.Forms.ComboBox
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents dbcOrigen1 As System.Windows.Forms.ComboBox
    Public WithEvents dbcVFamilia As System.Windows.Forms.ComboBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_8 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_7 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_6 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_4 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_3 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Frame3 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
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

    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button
    Const C_NINGUNA As String = "[ Vacío ... ]"

    Sub Imprime()
        Dim rptInvHojaControlsinGrupo As New rptInvHojaControlsinGrupo
        Dim rptInvHojaControl As New rptInvHojaControl
        Dim tbCurrent As CrystalDecisions.CrystalReports.Engine.Table
        Dim tliCurrent As CrystalDecisions.Shared.TableLogOnInfo
        Dim pvNum As New CrystalDecisions.Shared.ParameterValues
        Dim pdvNum As New CrystalDecisions.Shared.ParameterDiscreteValue

        On Error GoTo Merr
        Dim aParam(5) As Object
        Dim aValues(5) As Object
        Dim FechaInicio As Date
        Dim FechaFin As Date
        Dim TextoAdicional As String
        Dim Encabezado As String
        Dim ConsultaGuardar As String
        Dim ConsultaReporte As String
        Dim mblnTRansaccion As Boolean
        Dim cSELECTRPT As String
        Dim cSELECTGUARDAR As String
        Dim cORDERBYRPT As String

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        If ValidaDatos() = False Then Exit Sub
        Encabezado = "Hoja de Control Para Inventario Fisico"
        cSELECTRPT = ""
        cORDERBYRPT = ""
        ConsultaGuardar = " FROM         dbo.Inventario I INNER JOIN " & "dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " & "dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " & "dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen  CROSS JOIN " & "dbo.ConfiguracionGeneral " & "WHERE     (I.CodAlmacen = " & intCodSucursal & ") " & "GROUP BY I.CodArticulo, A.CodArticulo, A.DescArticulo, A.CodGrupo, A.CodFamilia, A.CodLinea, A.CodSubLinea, A.CodMarca, A.CodUnidad, " & "O.DescAlmacenOrigen , Al.DescAlmacen, I.CodAlmacen, A.CodAlmacenOrigen, A.CodModelo, A.CostoReal,A.PrecioPubDolar, dbo.ConfiguracionGeneral.TasaIva, A.CodigoAnt "

        If chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked Then
            ConsultaReporte = " FROM         dbo.Inventario I INNER JOIN " & "dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " & "dbo.CatUnidades U ON A.CodUnidad = U.CodUnidad INNER JOIN " & "dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " & "dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen LEFT OUTER JOIN " & "dbo.CatGrupos Gr ON A.CodGrupo = Gr.CodGrupo LEFT OUTER JOIN " & "dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia AND Gr.CodGrupo = Fa.CodGrupo LEFT OUTER JOIN " & "dbo.CatLineas Li ON A.CodGrupo = Li.CodGrupo AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea AND Fa.CodGrupo = Li.CodGrupo AND " & "Fa.CodFamilia = Li.CodFamilia LEFT OUTER JOIN " & "dbo.CatSubLineas su ON A.CodGrupo = su.CodGrupo AND A.CodFamilia = su.CodFamilia AND A.CodLinea = su.CodLinea AND " & "A.CodSubLinea = su.CodSubLinea AND Li.CodGrupo = su.CodGrupo AND Li.CodFamilia = su.CodFamilia AND " & "Li.CodLinea = su.CodLinea LEFT OUTER JOIN " & "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca AND Gr.CodGrupo = Ma.CodGrupo LEFT OUTER JOIN " & "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " & "Ma.CodGrupo = Mo.CodGrupo And Ma.CodMarca = Mo.CodMarca   CROSS JOIN " & "dbo.ConfiguracionGeneral " & "GROUP BY I.CodAlmacen, I.CodArticulo, A.CodArticulo, A.DescArticulo, A.CodAlmacenOrigen, A.CodGrupo, A.CodFamilia, A.CodLinea, A.CodSubLinea, A.CodMarca, " & "A.CodModelo, A.CodUnidad, U.DescUnidad, O.DescAlmacenOrigen, Al.DescAlmacen, Gr.DescGrupo, Fa.DescFamilia, Li.DescLinea, su.DescSubLinea, " & "Ma.DescMarca , Mo.DescModelo, U.DescUnidad , dbo.ConfiguracionGeneral.NombreEmp,  " & "A.PrecioPubDolar, dbo.ConfiguracionGeneral.TasaIva, A.OrigenAnt, CodigoAnt  "

            cSELECTRPT = "SELECT     I.CodAlmacen, Al.DescAlmacen, I.CodArticulo, A.DescArticulo, A.CodAlmacenOrigen, O.DescAlmacenOrigen, SUM(I.ExistenciaInicial) AS ExistenciaInicial, " & "MAX(I.UltimoCostoDLL) AS UltimoCosto, SUM(I.Entradas) AS Entradas, SUM(I.Salidas) AS Salida, SUM(I.Apartados) AS Apartados, " & "SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) AS ExistenciaTeorica, A.CodGrupo, Gr.DescGrupo, A.CodFamilia, " & "Ltrim(Rtrim(Fa.DescFamilia)) as DescFamilia, A.CodLinea, Ltrim(Rtrim(Li.DescLinea)) as DescLinea, A.CodSubLinea, Ltrim(Rtrim(su.DescSubLinea)) as DescSubLinea, A.CodMarca, Ltrim(Rtrim(Ma.DescMarca)) as DescMarca, A.CodModelo, Ltrim(Rtrim(Mo.DescModelo)) as DescModelo, " & "A.CodUnidad , U.DescUnidad ,Ltrim(Rtrim( dbo.ConfiguracionGeneral.NombreEmp)) as NombreEmpresa, A.OrigenAnt, " & "Case A.CodigoAnt When 0 Then '' Else  cast(A.OrigenAnt as nvarchar) + '-' + right('00000'+  Cast(A.CodigoAnt as varchar),5) End  as CodigoAnterior , A.CodigoAnt    "

        Else
            ConsultaReporte = " FROM         dbo.Inventario I INNER JOIN " & "dbo.CatArticulos A ON I.CodArticulo = A.CodArticulo INNER JOIN " & "dbo.CatUnidades U ON A.CodUnidad = U.CodUnidad INNER JOIN " & "dbo.CatOrigen O ON A.CodAlmacenOrigen = O.CodAlmacenOrigen INNER JOIN " & "dbo.CatAlmacen Al ON I.CodAlmacen = Al.CodAlmacen LEFT OUTER JOIN " & "dbo.CatGrupos Gr ON A.CodGrupo = Gr.CodGrupo LEFT OUTER JOIN " & "dbo.CatFamilias Fa ON A.CodGrupo = Fa.CodGrupo AND A.CodFamilia = Fa.CodFamilia AND Gr.CodGrupo = Fa.CodGrupo LEFT OUTER JOIN " & "dbo.CatLineas Li ON A.CodGrupo = Li.CodGrupo AND A.CodFamilia = Li.CodFamilia AND A.CodLinea = Li.CodLinea AND Fa.CodGrupo = Li.CodGrupo AND " & "Fa.CodFamilia = Li.CodFamilia LEFT OUTER JOIN " & "dbo.CatSubLineas su ON A.CodGrupo = su.CodGrupo AND A.CodFamilia = su.CodFamilia AND A.CodLinea = su.CodLinea AND " & "A.CodSubLinea = su.CodSubLinea AND Li.CodGrupo = su.CodGrupo AND Li.CodFamilia = su.CodFamilia AND " & "Li.CodLinea = su.CodLinea LEFT OUTER JOIN " & "dbo.CatMarcas Ma ON A.CodGrupo = Ma.CodGrupo AND A.CodMarca = Ma.CodMarca AND Gr.CodGrupo = Ma.CodGrupo LEFT OUTER JOIN " & "dbo.CatModelos Mo ON A.CodGrupo = Mo.CodGrupo AND A.CodMarca = Mo.CodMarca AND A.CodModelo = Mo.CodModelo AND " & "Ma.CodGrupo = Mo.CodGrupo And Ma.CodMarca = Mo.CodMarca   CROSS JOIN " & "dbo.ConfiguracionGeneral " & "GROUP BY I.CodAlmacen, I.CodArticulo, A.CodArticulo, A.DescArticulo, A.CodAlmacenOrigen,  A.CodFamilia, A.CodLinea, A.CodSubLinea, A.CodMarca, " & "A.CodModelo, A.CodUnidad, U.DescUnidad, O.DescAlmacenOrigen, Al.DescAlmacen, Gr.DescGrupo, Fa.DescFamilia, Li.DescLinea, su.DescSubLinea, " & "Ma.DescMarca , Mo.DescModelo, U.DescUnidad , dbo.ConfiguracionGeneral.NombreEmp,  " & "A.PrecioPubDolar, dbo.ConfiguracionGeneral.TasaIva, A.OrigenAnt, CodigoAnt , A.CodGrupo "

            cSELECTRPT = "SELECT     I.CodAlmacen, Al.DescAlmacen, I.CodArticulo, A.DescArticulo, A.CodAlmacenOrigen, O.DescAlmacenOrigen, SUM(I.ExistenciaInicial) AS ExistenciaInicial, " & "MAX(I.UltimoCostoDLL) AS UltimoCosto, SUM(I.Entradas) AS Entradas, SUM(I.Salidas) AS Salida, SUM(I.Apartados) AS Apartados, " & "SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) AS ExistenciaTeorica, A.CodFamilia, " & "Ltrim(Rtrim(Fa.DescFamilia)) as DescFamilia, A.CodLinea, Ltrim(Rtrim(Li.DescLinea)) as DescLinea, A.CodSubLinea, Ltrim(Rtrim(su.DescSubLinea)) as DescSubLinea, A.CodMarca, Ltrim(Rtrim(Ma.DescMarca)) as DescMarca, A.CodModelo, Ltrim(Rtrim(Mo.DescModelo)) as DescModelo, " & "A.CodUnidad , U.DescUnidad ,Ltrim(Rtrim( dbo.ConfiguracionGeneral.NombreEmp)) as NombreEmpresa, " & "Case A.CodigoAnt When 0 Then '' Else  cast(A.OrigenAnt as nvarchar) + '-' + right('00000'+  Cast(A.CodigoAnt as varchar),5) End  as CodigoAnterior , A.CodigoAnt  "
        End If

        cSELECTGUARDAR = "SELECT     I.CodAlmacen, A.CodAlmacenOrigen, A.CodGrupo, A.CodFamilia, A.CodLinea, A.CodSubLinea, A.CodMarca, A.CodModelo, A.CodArticulo, " & "SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) AS ExistenciaTeorica, NULL as ExistenciaFisica, 0 as Ajuste , A.CostoReal AS CostoUnitario "

        ConsultaGuardar = cSELECTGUARDAR & ConsultaGuardar

        Cnn.BeginTrans()
        mblnTRansaccion = True
        ConsultaGuardar = ConsultaGuardar & DevuelveQuery(1)
        ModStoredProcedures.PR_InvHojadeControl(ConsultaGuardar, txtCodSucursal.Text, IIf((Trim(txtCodOrigen.Text) = ""), -1, txtCodOrigen.Text), CStr(0), CStr(0), CStr(0), C_INSERCION, CStr(0))
        Cmd.Execute()

        Cnn.CommitTrans()
        mblnTRansaccion = False
        gStrSql = cSELECTRPT & ConsultaReporte & DevuelveQuery(2)

        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.UP_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        frmReportes.rsReport = Cmd.Execute
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        If frmReportes.rsReport.RecordCount = 0 Then
            MsgBox("No existe que reportar", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            Exit Sub
        Else
            rptInvHojaControlsinGrupo.SetDataSource(frmReportes.rsReport)
            rptInvHojaControl.SetDataSource(frmReportes.rsReport)
        End If

        'aParam(1) = "EncabezadoReporte"
        'aValues(1) = Encabezado
        'aParam(2) = "CodOrigen"
        'aValues(2) = CDbl(Numerico(txtCodOrigen.Text))
        'aParam(3) = "DescOrigen"
        'aValues(3) = Trim(dbcOrigen.Text)
        'aParam(4) = "IncluirExistencia"
        'aValues(4) = IIf((chkIncluirExistenciaTeorica.CheckState = System.Windows.Forms.CheckState.Checked), True, False)
        'aParam(5) = "OrdernarPorGrupo"
        'aValues(5) = IIf((chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked), True, False)

        'frmReportes.Report = IIf((chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked), rptInvHojaControl, rptInvHojaControlsinGrupo) 'Es el nombre del archivo que se incluyó en el proyecto
        'frmReportes.Imprime(Me.Text, aParam, aValues)


        If (chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked) Then

            If (Encabezado <> Nothing) Then
                pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
            End If

            If (CDbl(Numerico(txtCodOrigen.Text)) <> Nothing) Then
                pdvNum.Value = CDbl(Numerico(txtCodOrigen.Text)) : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("CodOrigen").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("CodOrigen").ApplyCurrentValues(pvNum)
            End If


            If (Trim(dbcOrigen1.Text) <> Nothing) Then
                pdvNum.Value = Trim(dbcOrigen1.Text) : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("DescOrigen").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("DescOrigen").ApplyCurrentValues(pvNum)
            End If


            If (chkIncluirExistenciaTeorica.CheckState = System.Windows.Forms.CheckState.Checked Or chkIncluirExistenciaTeorica.CheckState = System.Windows.Forms.CheckState.Unchecked <> Nothing) Then
                pdvNum.Value = IIf((chkIncluirExistenciaTeorica.CheckState = System.Windows.Forms.CheckState.Checked), True, False) : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("IncluirExistencia").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("IncluirExistencia").ApplyCurrentValues(pvNum)
            End If


            If (chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked Or chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Unchecked <> Nothing) Then
                pdvNum.Value = IIf((chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked), True, False) : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("OrdernarPorGrupo").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControl.DataDefinition.ParameterFields("OrdernarPorGrupo").ApplyCurrentValues(pvNum)
            End If

            frmReportes.reporteActual = rptInvHojaControl
            frmReportes.Show()


        Else

            If (Encabezado <> Nothing) Then
                pdvNum.Value = Encabezado : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("EncabezadoReporte").ApplyCurrentValues(pvNum)
            End If

            If (CDbl(Numerico(txtCodOrigen.Text)) <> Nothing) Then
                pdvNum.Value = CDbl(Numerico(txtCodOrigen.Text)) : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("CodOrigen").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("CodOrigen").ApplyCurrentValues(pvNum)
            End If


            If (Trim(dbcOrigen1.Text) <> Nothing) Then
                pdvNum.Value = Trim(dbcOrigen1.Text) : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("DescOrigen").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("DescOrigen").ApplyCurrentValues(pvNum)
            End If


            If (chkIncluirExistenciaTeorica.CheckState = System.Windows.Forms.CheckState.Checked Or chkIncluirExistenciaTeorica.CheckState = System.Windows.Forms.CheckState.Unchecked <> Nothing) Then
                pdvNum.Value = IIf((chkIncluirExistenciaTeorica.CheckState = System.Windows.Forms.CheckState.Checked), True, False) : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("IncluirExistencia").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("IncluirExistencia").ApplyCurrentValues(pvNum)
            End If


            If (chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked Or chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Unchecked <> Nothing) Then
                pdvNum.Value = IIf((chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked), True, False) : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("OrdernarPorGrupo").ApplyCurrentValues(pvNum)
            Else
                pdvNum.Value = "" : pvNum.Add(pdvNum)
                rptInvHojaControlsinGrupo.DataDefinition.ParameterFields("OrdernarPorGrupo").ApplyCurrentValues(pvNum)
            End If

            'frmReportes.reporteActual = IIf((chkOrdenarporGrupo.CheckState = System.Windows.Forms.CheckState.Checked), rptInvHojaControl, rptInvHojaControlsinGrupo) 'Es el nombre del archivo que se incluyó en el proyecto         
            frmReportes.reporteActual = rptInvHojaControlsinGrupo
            frmReportes.Show()
        End If
         
        'Exit Sub

Merr:
        If mblnTRansaccion = True Then Cnn.RollbackTrans()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Function ValidaDatos() As Boolean
        If CDbl(Numerico(txtCodSucursal.Text)) = 0 Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            MsgBox("Proporcione el almacén para generar el reporte.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            dbcSucursales.Focus()
            Exit Function
        End If
        If chkJoyeria.CheckState = System.Windows.Forms.CheckState.Unchecked And chkRelojeria.CheckState = System.Windows.Forms.CheckState.Unchecked And chkVarios.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            MsgBox("Debe seleccionar  un Grupo de Artículos para generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
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
        'If MDIMenuPrincipalCorpo.ActiveMdiChild.Name <> Me.Name Then
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
        'If MDIMenuPrincipalCorpo.ActiveMdiChild.Name <> Me.Name Then
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
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcJLinea.Text) = "" Then Me.dbcJLinea.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub dbcJSubLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcJSubLinea.Leave

        Dim Aux As Integer
        'If MDIMenuPrincipalCorpo.ActiveMdiChild.Name <> Me.Name Then
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
        dbcOrigen1_Leave(dbcOrigen1, New System.EventArgs())
MError:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub dbcOrigen1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen1.Enter
        Pon_Tool()
        gStrSql = "SELECT codAlmacenOrigen, RTrim(LTrim(descAlmacenOrigen)) as descAlmacenOrigen FROM CatOrigen ORDER BY CodAlmacenOrigen"
        ModDCombo.DCGotFocus(gStrSql, (Me.dbcOrigen1))
    End Sub

    Private Sub dbcOrigen1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcOrigen1.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            dbcSucursales.Focus()
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcOrigen1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcOrigen1.Leave
        Dim I As Integer
        'If MDIMenuPrincipalCorpo.ActiveMdiChild.Name <> Me.Name Then
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
        'If MDIMenuPrincipalCorpo.ActiveMdiChild.Name <> Me.Name Then
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
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcRMarca.Text) = "" Then Me.dbcRMarca.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    '''Relojeria --Modelos
    Private Sub dbcRmodelo_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcRModelo.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        ' If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcRModelo.Name Then Exit Sub
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
        'If MenuPrincipal.ActiveMDIChild.Name <> Me.Name Then
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
            mblnSalir = True
            Me.Close()
            eventSender.KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcSucursales_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyUp
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        mblnFueraChange = True
        If intCodSucursal = 0 Then
            txtCodSucursal.Text = ""
        Else
            txtCodSucursal.Text = CStr(intCodSucursal)
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        mblnFueraChange = True
        If intCodSucursal = 0 Then
            txtCodSucursal.Text = ""
        Else
            txtCodSucursal.Text = CStr(intCodSucursal)
        End If
        mblnFueraChange = False
    End Sub

    Private Sub dbcSucursales_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As MouseEventArgs) Handles dbcSucursales.MouseUp
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        mblnFueraChange = True
        If intCodSucursal = 0 Then
            txtCodSucursal.Text = ""
        Else
            txtCodSucursal.Text = CStr(intCodSucursal)
        End If
    End Sub

    Private Sub frminvHojadecontrol_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frminvHojadecontrol_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub Form_Initialize_Renamed()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
    End Sub

    Private Sub frminvHojadecontrol_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frminvHojadecontrol_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Sub Nuevo()

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
        mblnFueraChange = True
        txtCodOrigen.Text = ""
        dbcOrigen1.Text = ""
        dbcSucursales.Text = ""
        mblnFueraChange = False
        optCodigo.Checked = True
        chkIncluirExistenciaTeorica.CheckState = System.Windows.Forms.CheckState.Checked
        chkExistenciaMayorCero.CheckState = System.Windows.Forms.CheckState.Checked
    End Sub

    Private Sub frminvHojadecontrol_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Icono(Me, MDIMenuPrincipalCorpo)
        ModEstandar.CentrarForma(Me)
        Nuevo()
        '    txtCodSucursal = gintCodAlmacen
        '    txtCodSucursal_LostFocus
    End Sub

    Private Sub frminvHojadecontrol_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frminvHojadecontrol_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '                              Nuevo     Guardar      Cancelar      Eliminar    Buscar       Imprimir     Cerrar
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Sub Limpiar()
        Nuevo()
        txtCodSucursal.Text = ""
        dbcSucursales.Focus()
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
        '    If sstGrupos.Tab = 0 Then
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODJOYERIA & " ORDER BY DescFamilia"
        '    Else
        '        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " ORDER BY DescFamilia"
        '    End If
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
        '    Else
        '        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & Numerico(GridActivo.TextMatrix(GridActivo.Row, C_ColJCODFAMILIA)) & ")  ORDER BY DescLinea"
        '    End If
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
        If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcJSubLinea.Name Then Exit Sub
        gStrSql = "SELECT CodSubLinea,DescSubLinea=Ltrim(Rtrim(DescSubLinea)) From dbo.CatSubLineas Where (CodGrupo = " & gCODJOYERIA & ") And (CodFamilia = " & mintJFamilia & ")  And (CodLinea = " & mintJLinea & ") and DescSubLinea LIKE '" & Trim(dbcJSubLinea.Text) & "%' ORDER BY DescSubLinea"
        ModDCombo.DCChange(gStrSql, tecla)
        'LimpiaDatosPrecioYDescuento
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
            dbcVFamilia_Leave(dbcVFamilia, New System.EventArgs())
            '        KeyCode = 0
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcVFamilia_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcVFamilia.Name Then Exit Sub
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " and DescFamilia LIKE '" & Trim(dbcVFamilia.Text) & "%' ORDER BY DescFamilia"
        ModDCombo.DCChange(gStrSql, tecla)
        If Trim(Me.dbcVFamilia.Text) = "" Then
            dbcVFamilia_Leave(dbcVFamilia, New System.EventArgs())
        End If
    End Sub

    Private Sub dbcVFamilia_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.Enter
        Pon_Tool()
        gStrSql = "SELECT CodFamilia , DescFamilia =ltrim(rtrim(DescFamilia))  From CatFamilias Where CodGRupo = " & gCODVARIOS & " ORDER BY DescFamilia"
        ModDCombo.DCGotFocus(gStrSql, dbcVFamilia)
    End Sub

    Private Sub dbcVFamilia_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVFamilia.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If MDIMenuPrincipalCorpo.ActiveMdiChild.Name <> Me.Name Then
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
            End If
        End If
        mblnFueraChange = True
        If Trim(Me.dbcVFamilia.Text) = "" Then Me.dbcVFamilia.Text = C_TODAS
        mblnFueraChange = False
    End Sub

    Private Sub dbcVLinea_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcVLinea.KeyDown
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            Me.dbcVFamilia.Focus()
            eventSender.KeyCode = 0
            '    ElseIf KeyCode = vbKeyReturn Then
            '        dbcOrigen.SetFocus
        End If
        tecla = eventArgs.KeyCode
    End Sub

    Private Sub dbcVLinea_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.CursorChanged
        If mblnFueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> Me.dbcVLinea.Name Then Exit Sub
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & mintVFamilia & ") and DescLinea LIKE '" & Trim(dbcVLinea.Text) & "%' ORDER BY DescLinea"
        ModDCombo.DCChange(gStrSql, tecla)
        If Trim(Me.dbcVLinea.Text) = "" Then
            dbcVLinea_Leave(dbcVLinea, New System.EventArgs())
        End If
    End Sub

    Private Sub dbcVLinea_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.Enter
        If mblnFueraChange = True Then Exit Sub
        gStrSql = "SELECT CodLinea,DescLinea=Ltrim(Rtrim(DescLinea)) From dbo.CatLineas Where (CodGrupo = " & gCODVARIOS & ") And (CodFamilia = " & mintVFamilia & ")  ORDER BY DescLinea"
        ModDCombo.DCGotFocus(gStrSql, dbcVLinea)
    End Sub

    Private Sub dbcVLinea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcVLinea.Leave
        Dim Aux As Integer
        'If MDIMenuPrincipalCorpo.ActiveMdiChild.Name <> Me.Name Then
        '    Exit Sub
        'End If
        '    mblnFueraChange = True
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

    Private Sub optCodigo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCodigo.CheckedChanged
        If eventSender.Checked Then
            fraCodigo.Enabled = True
        End If
    End Sub

    Private Sub optDescripcion_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDescripcion.CheckedChanged
        If eventSender.Checked Then
            fraCodigo.Enabled = False
            'optCodActual.Enabled = False
            'optCodAnterior.Enabled=False
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

    Private Sub txtCodorigen_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCodOrigen.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSalir = True
            Me.Close()
        End If
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
            MsgBox("Código de Sucursal no existe." & vbNewLine & "Verifique Por Favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
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
            MsgBox("Código de Origen no existe." & vbNewLine & "Verifique Por Favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrCorpoNOMBREEMPRESA)
            txtCodOrigen.Text = ""
            dbcOrigen1.Focus()
        End If
        Exit Sub
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    '    Public Function DevuelveQuery() As String
    '        On Local Error GoTo MErr
    '        Dim I As Long
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

    '        'Obtener los códigos que va a tomar en cuenta en la consulta; estos códigos se enviarán como parámetros al
    '        'procedimiento almacenado que recopilará los datos

    '        nJOYERIA = Me.chkJoyeria.Value
    '        nRELOJERIA = Me.chkRelojeria.Value
    '        nVARIOS = Me.chkVarios.Value

    '        If nJOYERIA = 0 And nRELOJERIA = 0 And nVARIOS = 0 Then
    '            MsgBox "Debe elegir, por lo menos, un grupo con el cual generar el reporte", vbOKOnly + vbInformation, gstrCorpoNOMBREEMPRESA
    '            Exit Function
    '        End If

    '        cWHERE = " Having  "
    '        cORDERBY = " Order By "
    '        Select Case True
    '            Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS > 0
    '                'Todos los grupos
    '                cWHERE = cWHERE & " A.CodGrupo In (" & gCODJOYERIA & ", " & gCODRELOJERIA & ", " & gCODVARIOS & ") "
    '                Select Case True
    '                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        ' Todos
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
    '                        'Todos
    '                        cWHERE = cWHERE & " or (A.CodMarca <> " & 0 & ")"
    '                    Case mintRMarca > 0 And mintRModelo <= 0
    '                        cWHERE = cWHERE & " or (A.CodMarca = " & mintRMarca & ")"
    '                    Case mintRMarca > 0 And mintRModelo > 0
    '                        cWHERE = cWHERE & " or (A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & ")"
    '                End Select
    '                Select Case True
    '                    Case mintVFamilia <= 0 And mintVLinea <= 0
    '                        'Todos
    '                        cWHERE = cWHERE & " or (A.CodFamilia <> 0 and A.CodSubLinea is NULL))"
    '                    Case mintVFamilia > 0 And mintVLinea <= 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL))"
    '                    Case mintVFamilia > 0 And mintVLinea > 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL))"
    '                End Select
    '            Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS <= 0
    '                'Joyeria-Relojeria
    '                cWHERE = cWHERE & " A.CodGrupo <> " & gCODVARIOS
    '                Select Case True
    '                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        ' Todos
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
    '                        'Todos
    '                        cWHERE = cWHERE & " or (A.CodMarca <> " & 0 & "))"
    '                    Case mintRMarca > 0 And mintRModelo <= 0
    '                        cWHERE = cWHERE & " or (A.CodMarca = " & mintRMarca & "))"
    '                    Case mintRMarca > 0 And mintRModelo > 0
    '                        cWHERE = cWHERE & " or (CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & "))"
    '                End Select
    '            Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS > 0
    '                'Joyeria-Varios
    '                cWHERE = cWHERE & " A.CodGrupo <> " & gCODRELOJERIA
    '                Select Case True
    '                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        ' Todos
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
    '                        'Todos
    '                        cWHERE = cWHERE & " or (A.CodFamilia <> 0) and A.CodSubLinea is NULL)"
    '                    Case mintVFamilia > 0 And mintVLinea <= 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL))"
    '                    Case mintVFamilia > 0 And mintVLinea > 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL))"
    '                End Select
    '            Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS <= 0
    '                'Joyeria
    '                cWHERE = cWHERE & " A.CodGrupo = " & gCODJOYERIA
    '                Select Case True
    '                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        ' Todos
    '                        cWHERE = cWHERE & " and A.CodFamilia <> " & 0 & " and A.CodSubLinea is NOT NULL "
    '                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintJFamilia & " and A.CodSubLinea is NOT NULL "
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea is NOT NULL"
    '                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea
    '                End Select
    '            Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS > 0
    '                'Relojeria-Varios
    '                cWHERE = cWHERE & " A.CodGrupo <> " & gCODJOYERIA
    '                Select Case True
    '                    Case mintRMarca <= 0 And mintRModelo <= 0
    '                        'Todos
    '                        cWHERE = cWHERE & " and ((A.CodMarca <> " & 0 & ")"
    '                    Case mintRMarca > 0 And mintRModelo <= 0
    '                        cWHERE = cWHERE & " and ((A.CodMarca = " & mintRMarca & ")"
    '                    Case mintRMarca > 0 And mintRModelo > 0
    '                        cWHERE = cWHERE & " and ((A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & ")"
    '                End Select
    '                Select Case True
    '                    Case mintVFamilia <= 0 And mintVLinea <= 0
    '                        'Todos
    '                        cWHERE = cWHERE & " or (A.CodFamilia <> 0) and A.CodSubLinea is NULL)"
    '                    Case mintVFamilia > 0 And mintVLinea <= 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL))"
    '                    Case mintVFamilia > 0 And mintVLinea > 0
    '                        cWHERE = cWHERE & " or (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL))"
    '                End Select
    '            Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS <= 0
    '                'Relojeria
    '                cWHERE = cWHERE & " A.CodGrupo = " & gCODRELOJERIA
    '                Select Case True
    '                    Case mintRMarca <= 0 And mintRModelo <= 0
    '                        'Todos
    '                        cWHERE = cWHERE & " and A.CodMarca <> " & 0
    '                    Case mintRMarca > 0 And mintRModelo <= 0
    '                        cWHERE = cWHERE & " and A.CodMarca = " & mintRMarca
    '                    Case mintRMarca > 0 And mintRModelo > 0
    '                        cWHERE = cWHERE & " and A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo
    '                End Select
    '            Case nJOYERIA <= 0 And nRELOJERIA <= 0 And nVARIOS > 0
    '                'Varios
    '                cWHERE = cWHERE & " A.CodGrupo = " & gCODVARIOS
    '                Select Case True
    '                    Case mintVFamilia <= 0 And mintVLinea <= 0
    '                        'Todos
    '                        cWHERE = cWHERE & " and A.CodFamilia <> 0 and A.CodSubLinea is NULL "
    '                    Case mintVFamilia > 0 And mintVLinea <= 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintVFamilia & " and A.CodSubLinea is NULL "
    '                    Case mintVFamilia > 0 And mintVLinea > 0
    '                        cWHERE = cWHERE & " and A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & " and A.CodSubLinea is NULL "
    '                End Select
    '        End Select


    '        DevuelveQuery = cWHERE & " and I.CodAlmacen = " & txtCodSucursal
    '        If Trim(dbcOrigen) <> "" Then
    '            DevuelveQuery = DevuelveQuery & " And I.CodAlmacenOrigen = " & txtCodOrigen
    '        End If

    '        If optCodigo = True Then
    '            If optCodActual = True Then
    '                cORDERBY = cORDERBY + " I.CodArticulo "
    '            Else
    '                cORDERBY = cORDERBY + " a.CodigoAnt"
    '            End If
    '        Else
    '            cORDERBY = cORDERBY + " A.DescArticulo "
    '        End If

    '        DevuelveQuery = DevuelveQuery & cORDERBY

    '        Exit Function

    'MErr:
    '        If Err.Number <> 0 Then ModEstandar.MostrarError()
    '        '''
    '    End Function

    Public Function DevuelveQuery(ByRef lTipo As Integer) As String
        On Error GoTo Merr
        Dim I As Integer
        Dim cSELECT As String
        Dim cFROM As String
        Dim cWHERE As String
        Dim cGROUPBY As String
        Dim cORDERBY As String
        Dim rsLocal As ADODB.Recordset
        Dim cMSG As String

        Dim nJOYERIA As Integer
        Dim nRELOJERIA As Integer
        Dim nVARIOS As Integer
        Dim lExistenciaMayorCero As String

        'Obtener los códigos que va a tomar en cuenta en la consulta; estos códigos se enviarán como parámetros al
        'procedimiento almacenado que recopilará los datos
        nJOYERIA = Me.chkJoyeria.CheckState
        nRELOJERIA = Me.chkRelojeria.CheckState
        nVARIOS = Me.chkVarios.CheckState

        If nJOYERIA = 0 And nRELOJERIA = 0 And nVARIOS = 0 Then
            MsgBox("Debe elegir, por lo menos, un grupo con el cual generar el reporte", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrCorpoNOMBREEMPRESA)
            Exit Function
        End If

        cWHERE = " Having ( "
        cORDERBY = " Order By "
        Select Case True
            Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS > 0
                '''JOYERIA-RELOJERIA-VARIOS
                cWHERE = cWHERE & " ((A.CodGrupo = " & gCODJOYERIA & ") "
                Select Case True
                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        ' Todos
                        cWHERE = cWHERE & " and (A.CodFamilia <> 0)) "
                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & ")) "
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & ")) "
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & ")) "
                End Select

                cWHERE = cWHERE & " OR ((A.CodGrupo = " & gCODRELOJERIA & ") "
                Select Case True
                    Case mintRMarca <= 0 And mintRModelo <= 0
                        'Todos
                        cWHERE = cWHERE & " And (A.CodMarca <> 0)) "
                    Case mintRMarca > 0 And mintRModelo <= 0
                        cWHERE = cWHERE & " And (A.CodMarca = " & mintRMarca & ")) "
                    Case mintRMarca > 0 And mintRModelo > 0
                        cWHERE = cWHERE & " And (A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & ")) "
                End Select

                cWHERE = cWHERE & " OR ((A.CodGrupo = " & gCODVARIOS & ") "
                Select Case True
                    Case mintVFamilia <= 0 And mintVLinea <= 0
                        'Todos
                        cWHERE = cWHERE & " And (A.CodFamilia <> 0)) "
                    Case mintVFamilia > 0 And mintVLinea <= 0
                        cWHERE = cWHERE & " And (A.CodFamilia = " & mintVFamilia & ")) "
                    Case mintVFamilia > 0 And mintVLinea > 0
                        cWHERE = cWHERE & " And (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & ")) "
                End Select

            Case nJOYERIA > 0 And nRELOJERIA > 0 And nVARIOS <= 0
                '''Joyeria-Relojeria
                cWHERE = cWHERE & " ((A.CodGrupo = " & gCODJOYERIA & ") "
                Select Case True
                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        ' Todos
                        cWHERE = cWHERE & " and (A.CodFamilia <> " & 0 & ")) "
                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & " ))"
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " ))"
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & "))"
                End Select

                cWHERE = cWHERE & " OR ((A.CodGrupo = " & gCODRELOJERIA & ") "
                Select Case True
                    Case mintRMarca <= 0 And mintRModelo <= 0
                        'Todos
                        cWHERE = cWHERE & " And (A.CodMarca <> " & 0 & ")) "
                    Case mintRMarca > 0 And mintRModelo <= 0
                        cWHERE = cWHERE & " And (A.CodMarca = " & mintRMarca & "))"
                    Case mintRMarca > 0 And mintRModelo > 0
                        cWHERE = cWHERE & " And (A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & "))"
                End Select

            Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS > 0
                '''Joyeria-Varios
                cWHERE = cWHERE & " ((A.CodGrupo = " & gCODJOYERIA & ") "
                Select Case True
                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        'Todos
                        cWHERE = cWHERE & " and (A.CodFamilia <> 0)) "
                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & ")) "
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & ")) "
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & ")) "
                End Select

                cWHERE = cWHERE & " OR ((A.CodGrupo = " & gCODVARIOS & ") "
                Select Case True
                    Case mintVFamilia <= 0 And mintVLinea <= 0
                        'Todos
                        cWHERE = cWHERE & " And (A.CodFamilia <> 0)) "
                    Case mintVFamilia > 0 And mintVLinea <= 0
                        cWHERE = cWHERE & " And (A.CodFamilia = " & mintVFamilia & ")) "
                    Case mintVFamilia > 0 And mintVLinea > 0
                        cWHERE = cWHERE & " And (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & ")) "
                End Select

            Case nJOYERIA > 0 And nRELOJERIA <= 0 And nVARIOS <= 0
                '''Joyeria
                cWHERE = cWHERE & " ((A.CodGrupo = " & gCODJOYERIA & ") "
                Select Case True
                    Case mintJFamilia <= 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        ' Todos
                        cWHERE = cWHERE & " and (A.CodFamilia <> 0)) "
                    Case mintJFamilia > 0 And mintJLinea <= 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & ")) "
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea <= 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & ")) "
                    Case mintJFamilia > 0 And mintJLinea > 0 And mintJSubLinea > 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintJFamilia & " and A.CodLinea = " & mintJLinea & " and A.CodSubLinea = " & mintJSubLinea & ")) "
                End Select

            Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS > 0
                '''Relojeria-Varios
                cWHERE = cWHERE & " ((A.CodGrupo = " & gCODRELOJERIA & ") "
                Select Case True
                    Case mintRMarca <= 0 And mintRModelo <= 0
                        'Todos
                        cWHERE = cWHERE & " and (A.CodMarca <> 0)) "
                    Case mintRMarca > 0 And mintRModelo <= 0
                        cWHERE = cWHERE & " and (A.CodMarca = " & mintRMarca & ")) "
                    Case mintRMarca > 0 And mintRModelo > 0
                        cWHERE = cWHERE & " and (A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & ")) "
                End Select

                cWHERE = cWHERE & " OR ((A.CodGrupo = " & gCODVARIOS & ") "
                Select Case True
                    Case mintVFamilia <= 0 And mintVLinea <= 0
                        'Todos
                        cWHERE = cWHERE & " And (A.CodFamilia <> 0))"
                    Case mintVFamilia > 0 And mintVLinea <= 0
                        cWHERE = cWHERE & " And (A.CodFamilia = " & mintVFamilia & " ))"
                    Case mintVFamilia > 0 And mintVLinea > 0
                        cWHERE = cWHERE & " And (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & ")) "
                End Select

            Case nJOYERIA <= 0 And nRELOJERIA > 0 And nVARIOS <= 0
                '''Relojeria
                cWHERE = cWHERE & " ((A.CodGrupo = " & gCODRELOJERIA & " ) "
                Select Case True
                    Case mintRMarca <= 0 And mintRModelo <= 0
                        'Todos
                        cWHERE = cWHERE & " and (A.CodMarca <> 0 )) "
                    Case mintRMarca > 0 And mintRModelo <= 0
                        cWHERE = cWHERE & " and (A.CodMarca = " & mintRMarca & ")) "
                    Case mintRMarca > 0 And mintRModelo > 0
                        cWHERE = cWHERE & " and (A.CodMarca = " & mintRMarca & " and A.CodModelo = " & mintRModelo & ")) "
                End Select

            Case nJOYERIA <= 0 And nRELOJERIA <= 0 And nVARIOS > 0
                'Varios
                cWHERE = cWHERE & " ((A.CodGrupo = " & gCODVARIOS & ") "
                Select Case True
                    Case mintVFamilia <= 0 And mintVLinea <= 0
                        'Todos
                        cWHERE = cWHERE & " and (A.CodFamilia <> 0)) "
                    Case mintVFamilia > 0 And mintVLinea <= 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintVFamilia & ")) "
                    Case mintVFamilia > 0 And mintVLinea > 0
                        cWHERE = cWHERE & " and (A.CodFamilia = " & mintVFamilia & " and A.CodLinea = " & mintVLinea & ")) "
                End Select

        End Select

        DevuelveQuery = cWHERE & " ) And I.CodAlmacen = " & txtCodSucursal.Text
        If Trim(dbcOrigen1.Text) <> "" Then DevuelveQuery = DevuelveQuery & " And A.CodAlmacenOrigen = " & txtCodOrigen.Text

        lExistenciaMayorCero = " "
        If lTipo = 2 Then '''info para el rpt en CrystalR
            If chkExistenciaMayorCero.CheckState = System.Windows.Forms.CheckState.Checked Then
                lExistenciaMayorCero = "  And SUM(I.ExistenciaInicial) + SUM(I.Entradas) - SUM(I.Salidas) - SUM(I.Apartados) <> 0  "
            End If
        End If
        DevuelveQuery = DevuelveQuery & lExistenciaMayorCero

        If optCodigo.Checked = True Then
            If optCodActual.Checked = True Then
                cORDERBY = cORDERBY & " I.CodArticulo "
            Else
                cORDERBY = cORDERBY & " a.CodigoAnt "
            End If
        Else
            cORDERBY = cORDERBY & " A.DescArticulo "
        End If

        DevuelveQuery = DevuelveQuery & cORDERBY
        Exit Function

Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function


    Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.chkExistenciaMayorCero = New System.Windows.Forms.CheckBox()
        Me.chkIncluirExistenciaTeorica = New System.Windows.Forms.CheckBox()
        Me.chkRelojeria = New System.Windows.Forms.CheckBox()
        Me.chkVarios = New System.Windows.Forms.CheckBox()
        Me.chkJoyeria = New System.Windows.Forms.CheckBox()
        Me._Frame3_0 = New System.Windows.Forms.GroupBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.txtCodSucursal = New System.Windows.Forms.TextBox()
        Me.txtCodOrigen = New System.Windows.Forms.TextBox()
        Me._Frame3_1 = New System.Windows.Forms.GroupBox()
        Me.fraOrdenamiento = New System.Windows.Forms.GroupBox()
        Me.chkOrdenarporGrupo = New System.Windows.Forms.CheckBox()
        Me.optDescripcion = New System.Windows.Forms.RadioButton()
        Me.optCodigo = New System.Windows.Forms.RadioButton()
        Me.fraCodigo = New System.Windows.Forms.Panel()
        Me.optCodActual = New System.Windows.Forms.RadioButton()
        Me.optCodAnterior = New System.Windows.Forms.RadioButton()
        Me.dbcJFamilia = New System.Windows.Forms.ComboBox()
        Me.dbcJLinea = New System.Windows.Forms.ComboBox()
        Me.dbcJSubLinea = New System.Windows.Forms.ComboBox()
        Me.dbcVLinea = New System.Windows.Forms.ComboBox()
        Me.dbcRMarca = New System.Windows.Forms.ComboBox()
        Me.dbcRModelo = New System.Windows.Forms.ComboBox()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me.dbcOrigen1 = New System.Windows.Forms.ComboBox()
        Me.dbcVFamilia = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me._lblVentas_8 = New System.Windows.Forms.Label()
        Me._lblVentas_7 = New System.Windows.Forms.Label()
        Me._lblVentas_6 = New System.Windows.Forms.Label()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me._lblVentas_4 = New System.Windows.Forms.Label()
        Me._lblVentas_3 = New System.Windows.Forms.Label()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me.Frame3 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.fraOrdenamiento.SuspendLayout()
        Me.fraCodigo.SuspendLayout()
        CType(Me.Frame3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.chkExistenciaMayorCero)
        Me.Frame1.Controls.Add(Me.chkIncluirExistenciaTeorica)
        Me.Frame1.Controls.Add(Me.chkRelojeria)
        Me.Frame1.Controls.Add(Me.chkVarios)
        Me.Frame1.Controls.Add(Me.chkJoyeria)
        Me.Frame1.Controls.Add(Me._Frame3_0)
        Me.Frame1.Controls.Add(Me.Frame4)
        Me.Frame1.Controls.Add(Me.txtCodSucursal)
        Me.Frame1.Controls.Add(Me.txtCodOrigen)
        Me.Frame1.Controls.Add(Me._Frame3_1)
        Me.Frame1.Controls.Add(Me.fraOrdenamiento)
        Me.Frame1.Controls.Add(Me.dbcJFamilia)
        Me.Frame1.Controls.Add(Me.dbcJLinea)
        Me.Frame1.Controls.Add(Me.dbcJSubLinea)
        Me.Frame1.Controls.Add(Me.dbcVLinea)
        Me.Frame1.Controls.Add(Me.dbcRMarca)
        Me.Frame1.Controls.Add(Me.dbcRModelo)
        Me.Frame1.Controls.Add(Me.dbcSucursales)
        Me.Frame1.Controls.Add(Me.dbcOrigen1)
        Me.Frame1.Controls.Add(Me.dbcVFamilia)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me._lblVentas_8)
        Me.Frame1.Controls.Add(Me._lblVentas_7)
        Me.Frame1.Controls.Add(Me._lblVentas_6)
        Me.Frame1.Controls.Add(Me._lblVentas_5)
        Me.Frame1.Controls.Add(Me._lblVentas_4)
        Me.Frame1.Controls.Add(Me._lblVentas_3)
        Me.Frame1.Controls.Add(Me._lblVentas_0)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(441, 434)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'chkExistenciaMayorCero
        '
        Me.chkExistenciaMayorCero.BackColor = System.Drawing.SystemColors.Control
        Me.chkExistenciaMayorCero.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkExistenciaMayorCero.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkExistenciaMayorCero.Location = New System.Drawing.Point(297, 296)
        Me.chkExistenciaMayorCero.Name = "chkExistenciaMayorCero"
        Me.chkExistenciaMayorCero.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkExistenciaMayorCero.Size = New System.Drawing.Size(124, 17)
        Me.chkExistenciaMayorCero.TabIndex = 19
        Me.chkExistenciaMayorCero.Text = "Sólo con existencia"
        Me.chkExistenciaMayorCero.UseVisualStyleBackColor = False
        '
        'chkIncluirExistenciaTeorica
        '
        Me.chkIncluirExistenciaTeorica.BackColor = System.Drawing.SystemColors.Control
        Me.chkIncluirExistenciaTeorica.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncluirExistenciaTeorica.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIncluirExistenciaTeorica.Location = New System.Drawing.Point(20, 296)
        Me.chkIncluirExistenciaTeorica.Name = "chkIncluirExistenciaTeorica"
        Me.chkIncluirExistenciaTeorica.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncluirExistenciaTeorica.Size = New System.Drawing.Size(150, 18)
        Me.chkIncluirExistenciaTeorica.TabIndex = 18
        Me.chkIncluirExistenciaTeorica.Text = "Incluir existencia teórica"
        Me.chkIncluirExistenciaTeorica.UseVisualStyleBackColor = False
        '
        'chkRelojeria
        '
        Me.chkRelojeria.BackColor = System.Drawing.SystemColors.Control
        Me.chkRelojeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkRelojeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkRelojeria.Location = New System.Drawing.Point(20, 176)
        Me.chkRelojeria.Name = "chkRelojeria"
        Me.chkRelojeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkRelojeria.Size = New System.Drawing.Size(81, 17)
        Me.chkRelojeria.TabIndex = 9
        Me.chkRelojeria.Text = "Relojería"
        Me.chkRelojeria.UseVisualStyleBackColor = False
        '
        'chkVarios
        '
        Me.chkVarios.BackColor = System.Drawing.SystemColors.Control
        Me.chkVarios.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVarios.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkVarios.Location = New System.Drawing.Point(20, 240)
        Me.chkVarios.Name = "chkVarios"
        Me.chkVarios.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVarios.Size = New System.Drawing.Size(81, 17)
        Me.chkVarios.TabIndex = 10
        Me.chkVarios.Text = "Varios"
        Me.chkVarios.UseVisualStyleBackColor = False
        '
        'chkJoyeria
        '
        Me.chkJoyeria.BackColor = System.Drawing.SystemColors.Control
        Me.chkJoyeria.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkJoyeria.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkJoyeria.Location = New System.Drawing.Point(20, 88)
        Me.chkJoyeria.Name = "chkJoyeria"
        Me.chkJoyeria.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkJoyeria.Size = New System.Drawing.Size(81, 17)
        Me.chkJoyeria.TabIndex = 8
        Me.chkJoyeria.Text = "Joyería"
        Me.chkJoyeria.UseVisualStyleBackColor = False
        '
        '_Frame3_0
        '
        Me._Frame3_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame3_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame3_0.Location = New System.Drawing.Point(16, 160)
        Me._Frame3_0.Name = "_Frame3_0"
        Me._Frame3_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame3_0.Size = New System.Drawing.Size(417, 2)
        Me._Frame3_0.TabIndex = 7
        Me._Frame3_0.TabStop = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(16, 224)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(417, 2)
        Me.Frame4.TabIndex = 4
        Me.Frame4.TabStop = False
        '
        'txtCodSucursal
        '
        Me.txtCodSucursal.AcceptsReturn = True
        Me.txtCodSucursal.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodSucursal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodSucursal.Enabled = False
        Me.txtCodSucursal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodSucursal.Location = New System.Drawing.Point(92, 24)
        Me.txtCodSucursal.MaxLength = 0
        Me.txtCodSucursal.Name = "txtCodSucursal"
        Me.txtCodSucursal.ReadOnly = True
        Me.txtCodSucursal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodSucursal.Size = New System.Drawing.Size(49, 20)
        Me.txtCodSucursal.TabIndex = 2
        Me.txtCodSucursal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCodOrigen
        '
        Me.txtCodOrigen.AcceptsReturn = True
        Me.txtCodOrigen.BackColor = System.Drawing.SystemColors.Window
        Me.txtCodOrigen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCodOrigen.Enabled = False
        Me.txtCodOrigen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCodOrigen.Location = New System.Drawing.Point(92, 48)
        Me.txtCodOrigen.MaxLength = 4
        Me.txtCodOrigen.Name = "txtCodOrigen"
        Me.txtCodOrigen.ReadOnly = True
        Me.txtCodOrigen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCodOrigen.Size = New System.Drawing.Size(49, 20)
        Me.txtCodOrigen.TabIndex = 5
        Me.txtCodOrigen.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        '_Frame3_1
        '
        Me._Frame3_1.BackColor = System.Drawing.SystemColors.Control
        Me._Frame3_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame3_1.Location = New System.Drawing.Point(16, 80)
        Me._Frame3_1.Name = "_Frame3_1"
        Me._Frame3_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame3_1.Size = New System.Drawing.Size(417, 2)
        Me._Frame3_1.TabIndex = 1
        Me._Frame3_1.TabStop = False
        '
        'fraOrdenamiento
        '
        Me.fraOrdenamiento.BackColor = System.Drawing.SystemColors.Control
        Me.fraOrdenamiento.Controls.Add(Me.chkOrdenarporGrupo)
        Me.fraOrdenamiento.Controls.Add(Me.optDescripcion)
        Me.fraOrdenamiento.Controls.Add(Me.optCodigo)
        Me.fraOrdenamiento.Controls.Add(Me.fraCodigo)
        Me.fraOrdenamiento.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraOrdenamiento.Location = New System.Drawing.Point(8, 320)
        Me.fraOrdenamiento.Name = "fraOrdenamiento"
        Me.fraOrdenamiento.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraOrdenamiento.Size = New System.Drawing.Size(425, 99)
        Me.fraOrdenamiento.TabIndex = 21
        Me.fraOrdenamiento.TabStop = False
        Me.fraOrdenamiento.Text = " Ordenado por .... "
        '
        'chkOrdenarporGrupo
        '
        Me.chkOrdenarporGrupo.BackColor = System.Drawing.SystemColors.Control
        Me.chkOrdenarporGrupo.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkOrdenarporGrupo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkOrdenarporGrupo.Location = New System.Drawing.Point(8, 16)
        Me.chkOrdenarporGrupo.Name = "chkOrdenarporGrupo"
        Me.chkOrdenarporGrupo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkOrdenarporGrupo.Size = New System.Drawing.Size(65, 17)
        Me.chkOrdenarporGrupo.TabIndex = 22
        Me.chkOrdenarporGrupo.Text = "Grupo"
        Me.chkOrdenarporGrupo.UseVisualStyleBackColor = False
        '
        'optDescripcion
        '
        Me.optDescripcion.BackColor = System.Drawing.SystemColors.Control
        Me.optDescripcion.Cursor = System.Windows.Forms.Cursors.Default
        Me.optDescripcion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optDescripcion.Location = New System.Drawing.Point(232, 24)
        Me.optDescripcion.Name = "optDescripcion"
        Me.optDescripcion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optDescripcion.Size = New System.Drawing.Size(110, 17)
        Me.optDescripcion.TabIndex = 28
        Me.optDescripcion.TabStop = True
        Me.optDescripcion.Text = "Por Descripción"
        Me.optDescripcion.UseVisualStyleBackColor = False
        '
        'optCodigo
        '
        Me.optCodigo.BackColor = System.Drawing.SystemColors.Control
        Me.optCodigo.Checked = True
        Me.optCodigo.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCodigo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCodigo.Location = New System.Drawing.Point(88, 24)
        Me.optCodigo.Name = "optCodigo"
        Me.optCodigo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCodigo.Size = New System.Drawing.Size(105, 17)
        Me.optCodigo.TabIndex = 24
        Me.optCodigo.TabStop = True
        Me.optCodigo.Text = "Por Código"
        Me.optCodigo.UseVisualStyleBackColor = False
        '
        'fraCodigo
        '
        Me.fraCodigo.BackColor = System.Drawing.SystemColors.Control
        Me.fraCodigo.Controls.Add(Me.optCodActual)
        Me.fraCodigo.Controls.Add(Me.optCodAnterior)
        Me.fraCodigo.Cursor = System.Windows.Forms.Cursors.Default
        Me.fraCodigo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraCodigo.Location = New System.Drawing.Point(88, 44)
        Me.fraCodigo.Name = "fraCodigo"
        Me.fraCodigo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraCodigo.Size = New System.Drawing.Size(107, 46)
        Me.fraCodigo.TabIndex = 35
        '
        'optCodActual
        '
        Me.optCodActual.BackColor = System.Drawing.SystemColors.Control
        Me.optCodActual.Checked = True
        Me.optCodActual.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCodActual.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCodActual.Location = New System.Drawing.Point(3, 3)
        Me.optCodActual.Name = "optCodActual"
        Me.optCodActual.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCodActual.Size = New System.Drawing.Size(78, 21)
        Me.optCodActual.TabIndex = 25
        Me.optCodActual.TabStop = True
        Me.optCodActual.Text = "Actual"
        Me.optCodActual.UseVisualStyleBackColor = False
        '
        'optCodAnterior
        '
        Me.optCodAnterior.BackColor = System.Drawing.SystemColors.Control
        Me.optCodAnterior.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCodAnterior.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCodAnterior.Location = New System.Drawing.Point(3, 19)
        Me.optCodAnterior.Name = "optCodAnterior"
        Me.optCodAnterior.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCodAnterior.Size = New System.Drawing.Size(78, 24)
        Me.optCodAnterior.TabIndex = 26
        Me.optCodAnterior.TabStop = True
        Me.optCodAnterior.Text = "Anterior"
        Me.optCodAnterior.UseVisualStyleBackColor = False
        '
        'dbcJFamilia
        '
        Me.dbcJFamilia.Location = New System.Drawing.Point(178, 88)
        Me.dbcJFamilia.Name = "dbcJFamilia"
        Me.dbcJFamilia.Size = New System.Drawing.Size(253, 21)
        Me.dbcJFamilia.TabIndex = 11
        '
        'dbcJLinea
        '
        Me.dbcJLinea.Location = New System.Drawing.Point(178, 112)
        Me.dbcJLinea.Name = "dbcJLinea"
        Me.dbcJLinea.Size = New System.Drawing.Size(253, 21)
        Me.dbcJLinea.TabIndex = 12
        '
        'dbcJSubLinea
        '
        Me.dbcJSubLinea.Location = New System.Drawing.Point(178, 136)
        Me.dbcJSubLinea.Name = "dbcJSubLinea"
        Me.dbcJSubLinea.Size = New System.Drawing.Size(253, 21)
        Me.dbcJSubLinea.TabIndex = 13
        '
        'dbcVLinea
        '
        Me.dbcVLinea.Location = New System.Drawing.Point(178, 264)
        Me.dbcVLinea.Name = "dbcVLinea"
        Me.dbcVLinea.Size = New System.Drawing.Size(253, 21)
        Me.dbcVLinea.TabIndex = 17
        '
        'dbcRMarca
        '
        Me.dbcRMarca.Location = New System.Drawing.Point(178, 176)
        Me.dbcRMarca.Name = "dbcRMarca"
        Me.dbcRMarca.Size = New System.Drawing.Size(253, 21)
        Me.dbcRMarca.TabIndex = 14
        '
        'dbcRModelo
        '
        Me.dbcRModelo.Location = New System.Drawing.Point(178, 200)
        Me.dbcRModelo.Name = "dbcRModelo"
        Me.dbcRModelo.Size = New System.Drawing.Size(253, 21)
        Me.dbcRModelo.TabIndex = 15
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(156, 24)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(275, 21)
        Me.dbcSucursales.TabIndex = 3
        '
        'dbcOrigen1
        '
        Me.dbcOrigen1.Location = New System.Drawing.Point(156, 48)
        Me.dbcOrigen1.Name = "dbcOrigen1"
        Me.dbcOrigen1.Size = New System.Drawing.Size(275, 21)
        Me.dbcOrigen1.TabIndex = 6
        '
        'dbcVFamilia
        '
        Me.dbcVFamilia.Location = New System.Drawing.Point(178, 240)
        Me.dbcVFamilia.Name = "dbcVFamilia"
        Me.dbcVFamilia.Size = New System.Drawing.Size(253, 21)
        Me.dbcVFamilia.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(20, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(73, 17)
        Me.Label1.TabIndex = 34
        Me.Label1.Text = "Sucursal :"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(20, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "Origen :"
        '
        '_lblVentas_8
        '
        Me._lblVentas_8.AutoSize = True
        Me._lblVentas_8.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_8.Location = New System.Drawing.Point(116, 263)
        Me._lblVentas_8.Name = "_lblVentas_8"
        Me._lblVentas_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_8.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_8.TabIndex = 32
        Me._lblVentas_8.Text = "Línea"
        '
        '_lblVentas_7
        '
        Me._lblVentas_7.AutoSize = True
        Me._lblVentas_7.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_7.Location = New System.Drawing.Point(116, 240)
        Me._lblVentas_7.Name = "_lblVentas_7"
        Me._lblVentas_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_7.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_7.TabIndex = 31
        Me._lblVentas_7.Text = "Familia"
        '
        '_lblVentas_6
        '
        Me._lblVentas_6.AutoSize = True
        Me._lblVentas_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_6.Location = New System.Drawing.Point(116, 200)
        Me._lblVentas_6.Name = "_lblVentas_6"
        Me._lblVentas_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_6.Size = New System.Drawing.Size(42, 13)
        Me._lblVentas_6.TabIndex = 30
        Me._lblVentas_6.Text = "Modelo"
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.AutoSize = True
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_5.Location = New System.Drawing.Point(116, 176)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(37, 13)
        Me._lblVentas_5.TabIndex = 29
        Me._lblVentas_5.Text = "Marca"
        '
        '_lblVentas_4
        '
        Me._lblVentas_4.AutoSize = True
        Me._lblVentas_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_4.Location = New System.Drawing.Point(116, 138)
        Me._lblVentas_4.Name = "_lblVentas_4"
        Me._lblVentas_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_4.Size = New System.Drawing.Size(54, 13)
        Me._lblVentas_4.TabIndex = 27
        Me._lblVentas_4.Text = "SubLínea"
        '
        '_lblVentas_3
        '
        Me._lblVentas_3.AutoSize = True
        Me._lblVentas_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_3.Location = New System.Drawing.Point(116, 113)
        Me._lblVentas_3.Name = "_lblVentas_3"
        Me._lblVentas_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_3.Size = New System.Drawing.Size(35, 13)
        Me._lblVentas_3.TabIndex = 23
        Me._lblVentas_3.Text = "Línea"
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_0.Location = New System.Drawing.Point(116, 88)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_0.TabIndex = 20
        Me._lblVentas_0.Text = "Familia"
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 458)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 141
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(12, 458)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 140
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frminvHojadecontrol
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(456, 506)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(316, 147)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frminvHojadecontrol"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Impresión de la Hoja de Control"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.fraOrdenamiento.ResumeLayout(False)
        Me.fraCodigo.ResumeLayout(False)
        CType(Me.Frame3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub
End Class