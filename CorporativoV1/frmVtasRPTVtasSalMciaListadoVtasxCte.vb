Option Strict Off
Option Explicit On
Imports System.IO
Imports ADODB
Imports Microsoft.VisualBasic.Compatibility
Imports Microsoft.Office.Interop
Public Class frmVtasRPTVtasSalMciaListadoVtasxCte
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer
    ''' ***************************************************************************************************************************************************
    ''' LISTADO DE VENTAS AGRUPADO POR CLIENTE ORDENADO POR IMPORTE DESCENDENTE - SEGENERA DIRECTAMENTE A EXCEL - BASE
    ''' PARA ETIQUETAS DE CLIENTES - SE MODIFICO DESHABILITARMENU/ACCESOS/LLENAVFORMAS
    ''' NO OLVIDAR QUE ETIQUETA Y COMBO DEL CLIENTE ESTAN OCULTOS EN EL FRAME DE SUCURSAL
    ''' 06OCT2006 - MAVF
    '''
    ''' MODIFICACION DEL REPORTE.- SE AGREGARON CAMPOS NUEVOS AL REPORTE (TELCASA-TELOFICINA-EMAIL) E INFORMACION DE REPARACIONES
    ''' 11NOV2010 - MAVF Ver
    '''
    ''' SE AGREGO FORMATO TIPO TEXTO PARA COLUMNAS CP-TEL1-TEL2
    ''' 06DIC2010 - MAVF  Ver
    '''
    ''' Ver 1.1       Estatus: Aprobado
    ''' ***************************************************************************************************************************************************

    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkMostrarImporte As System.Windows.Forms.CheckBox
    Public WithEvents chkTodas As System.Windows.Forms.CheckBox
    Public WithEvents dbcSucursal As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_0 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_0 As System.Windows.Forms.GroupBox
    Public WithEvents dtpDesde As System.Windows.Forms.DateTimePicker
    Public WithEvents dtpHasta As System.Windows.Forms.DateTimePicker
    Public WithEvents _lblVentas_1 As System.Windows.Forms.Label
    Public WithEvents _lblVentas_2 As System.Windows.Forms.Label
    Public WithEvents _fraVtas_1 As System.Windows.Forms.GroupBox
    Public WithEvents chkImpuesto As System.Windows.Forms.CheckBox
    Public WithEvents txtMensaje As System.Windows.Forms.TextBox
    Public WithEvents dbcCliente As System.Windows.Forms.ComboBox
    Public WithEvents _lblVentas_5 As System.Windows.Forms.Label
    Public WithEvents _lblRpt_2 As System.Windows.Forms.Label
    Public WithEvents fraVtas As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents lblRpt As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblVentas As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Const C_TODAS As String = "[ Todas ... ]"
    Const C_TODOS As String = "[ Todos ... ]"
    Const C_NINGUNA As String = "[ Vacío ... ]"

    Dim msglTiempoCambioI As Single 'Variable para controlar el cambio en el date picker de fecha Inicial
    Dim msglTiempoCambioF As Single 'Variable para controlar el cambio en el date picker de fecha Final
    Dim mblnTecleoFechaI As Boolean
    Dim mblnTecleoFechaF As Boolean

    Dim mblnFueraChange As Boolean
    Dim mintCodSucursal As Integer
    Dim mintCodCliente As Integer
    Dim tecla As Integer
    Dim mblnSalir As Boolean

    Dim RsAux As ADODB.Recordset
    Dim ObjExcel As Object
    Dim objLibro As Excel.Workbook
    Dim objHoja As Excel.Worksheet
    Dim ColumSepar As Integer
    Dim ColumCtoL As Integer
    Dim cmd As ADODB.Command
    Public WithEvents btnNuevo As Button
    Public WithEvents btnImprimir As Button

    Const C_ENCABEZADO As Integer = 1

    Public Sub Limpiar()
        On Error Resume Next
        Call Me.Nuevo()
        Me.chkTodas.Focus()
    End Sub

    Public Sub Nuevo()
        Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        chkTodas_CheckStateChanged(chkTodas, New System.EventArgs())

        mblnFueraChange = True
        Me.dbcCliente.Text = C_TODOS
        Me.dbcCliente.Tag = ""
        mintCodCliente = 0
        mblnFueraChange = False

        Me.dtpDesde.Value = Format(Today, "dd/MMM/yyyy")
        Me.dtpHasta.Value = Format(Today, "dd/MMM/yyyy")
        Me.chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMostrarImporte.CheckState = System.Windows.Forms.CheckState.Checked
        Me.txtMensaje.Text = ""
        mblnTecleoFechaI = False
        mblnTecleoFechaF = False
    End Sub

    Function DevuelveQuery() As String
        On Error GoTo Err_Renamed
        Dim Sql As String

        Sql = ""
        Sql = Sql & "Select   A.CodCliente, B.DescCliente, B.Domicilio, B.Colonia, B.Ciudad, B.CP, "
        Sql = Sql & "         Sum(Tipo) as TipoC, Sum(A.Importe) as Impte, "
        Sql = Sql & "         B.TelCasa , B.TelOficina, B.Email "
        Sql = Sql & "From     ( "
        Sql = Sql & "         SELECT   CodCliente, Nombre, 0 as Tipo, "

        If chkImpuesto.CheckState = 1 Then
            Sql = Sql & "         sum(ROUND(PrecioReal * (Cantidad - CantidadDev) + CASE WHEN NumPartida = 1 THEN Redondeo ELSE 0 END,2)) AS Importe "
        Else
            Sql = Sql & "         Sum(ROUND((PrecioListaSinIva - Descuento) * (Cantidad - CantidadDev) + CASE WHEN NumPartida = 1 THEN Redondeo ELSE 0 END,2)) AS Importe"
        End If

        Sql = Sql & "         FROM  DBO.VTAS_SALIDAMCIA ('" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "', '" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "') "
        Sql = Sql & "         Where (Cantidad - CantidadDev) > 0 "

        If chkTodas.CheckState = 0 Then
            Sql = Sql & "         And   CodSucursal = " & mintCodSucursal & " "
        End If

        Sql = Sql & "         And     CodCliente <> 1 "
        Sql = Sql & "         Group   by CodCliente, Nombre "

        ''' 11NOV2010 - MAVF
        Sql = Sql & "         UNION "
        Sql = Sql & "         Select  CodCliente, Nombre, 1 as Tipo, 0 AS Importe "
        Sql = Sql & "         From    Reparaciones (Nolock) "
        Sql = Sql & "         where   FechaReparacion Between '" & Format(dtpDesde.Value, C_FORMATFECHAGUARDAR) & "' and '" & Format(dtpHasta.Value, C_FORMATFECHAGUARDAR) & "' "
        Sql = Sql & "         And     Estatus <> 'C' "
        If chkTodas.CheckState = 0 Then
            Sql = Sql & "         And     CodSucursal = " & mintCodSucursal & " "
        End If
        ''' *******************************************************************

        Sql = Sql & "            ) as A Inner Join CatClientes B (Nolock) On A.CodCliente = B.CodCliente "
        Sql = Sql & "Group by A.CodCliente, B.Domicilio, B.Colonia, B.Ciudad, B.CP, B.DescCliente, B.TelCasa, B.TelOficina, B.Email "
        Sql = Sql & "Order   by Impte Desc "

        DevuelveQuery = Sql

Err_Renamed:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Function


    Public Sub Imprime()

        On Error GoTo Merr
        Dim lStrSql As String

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        If Not ValidaDatos() Then
            Cursor = System.Windows.Forms.Cursors.Default
            Exit Sub
        End If

        lStrSql = DevuelveQuery()
        gStrSql = lStrSql
        ModEstandar.BorraCmd()
        cmd.CommandTimeout = 300
        cmd.CommandText = "dbo.UP_Select_Datos"
        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        cmd.Parameters.Append(cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        cmd.Parameters.Append(cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, gStrSql))
        RsGral = cmd.Execute

        If RsGral.RecordCount = 0 Then
            Cursor = System.Windows.Forms.Cursors.Default
            MsgBox("No existen datos para el rango de fechas indicado", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
        Else
            EnviaExcel()
        End If
        cmd.CommandTimeout = 90
        Cursor = System.Windows.Forms.Cursors.Default

Merr:
        If Err.Number <> 0 Then
            Cursor = System.Windows.Forms.Cursors.Default
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
                Cursor = System.Windows.Forms.Cursors.Default
                MsgBox("Si no quiere imprimir los resultados de todas las sucursales, seleccione una de ellas", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dbcSucursal.Focus()
            Case Me.dtpDesde.Value > Me.dtpHasta.Value
                Cursor = System.Windows.Forms.Cursors.Default
                MsgBox("La Fecha Inicial debe ser MENOR a la Fecha Límite", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                ValidaDatos = False
                Me.dtpDesde.Focus()
            Case Else
                ValidaDatos = True
        End Select
    End Function

    ''' *******************************************************************************************************************
    ''' FUNCIONES PARA GENERACION DE ARCHIVO A EXCEL **********************************************************************
    ''' 06OCT2006 - MAVF
    Sub EnviaExcel()
        On Error GoTo Err_Renamed
        Dim Archivo As String
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        System.Windows.Forms.Application.DoEvents()
        If Dir(gstrCorpoDriveLocal & "\Sistema\", FileAttribute.Directory + FileAttribute.Hidden) = "" Then
            MsgBox("No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            Exit Sub
        End If

        Archivo = "LV" & CStr(Format(Month(Today), "00")) & CStr(Format((Today), "00")) & (CStr(Format(Year(Today) + "00"))) & ".xls"
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", FileAttribute.Directory) = "" Then
            MkDir(gstrCorpoDriveLocal & "\Sistema\Informes\")
        End If
        If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo, FileAttribute.Archive) <> "" Then
            Kill(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo)
        End If

        ''' INSTANCIAS DE EXCEL PARA ARCHIVO
        ObjExcel = CreateObject("Excel.Application")
        objLibro = ObjExcel.Workbooks.Add
        objHoja = objLibro.ActiveSheet
        ObjExcel.Visible = False
        objLibro.Sheets(1).Select()
        objHoja = objLibro.ActiveSheet
        objLibro.ActiveSheet.Name = "Listado Ventas"

        Encabezado()
        LlenaDatos()

        objLibro.SaveAs(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        System.Windows.Forms.Application.DoEvents()
        Select Case MsgBox("Se ha creado el archivo " & Archivo & " ¿Desea abrirlo?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa)
            Case MsgBoxResult.Yes
                ObjExcel.Visible = True
                ObjExcel = Nothing
                objLibro = Nothing
                objHoja = Nothing
            Case MsgBoxResult.No Or MsgBoxResult.Cancel
                CierraInstanciasdeExcel(1)
        End Select

Err_Renamed:
        If Err.Number = 70 Then
            MsgBox("No se puede generar un nuevo archivo hasta que el anterior este cerrado.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            CierraInstanciasdeExcel(2)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        ElseIf Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    '''''' *******************************************************************************************************************
    '''''' FUNCIONES PARA GENERACION DE ARCHIVO A EXCEL **********************************************************************
    '''''' 06OCT2006 - MAVF
    '''Sub EnviaExcel()
    '''On Error GoTo Err
    '''   Dim Archivo As String
    '''
    '''   Screen.MousePointer = ccHourglass
    '''   DoEvents
    '''   If Dir(gstrCorpoDriveLocal & "\Sistema\", vbDirectory + vbHidden) = "" Then
    '''       MsgBox "No Existe la Carpeta Sistema, no se puede guardar el archivo, Favor de Verificar...", vbOKOnly + vbInformation, gstrNombCortoEmpresa
    '''       Exit Sub
    '''   End If
    '''
    '''   Archivo = "LV" & CStr(Format(Month(Date), "00")) & CStr(Format(Day(Date), "00")) & Right(CStr(Format(Year(Date), "00")), 2) & ".xls"
    '''   If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\", vbDirectory) = "" Then
    '''       MkDir gstrCorpoDriveLocal & "\Sistema\Informes\"
    '''   End If
    '''   If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo, vbArchive) <> "" Then
    '''       Kill gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo
    '''   End If
    '''
    '''   ''' INSTANCIAS DE EXCEL PARA ARCHIVO
    '''   Set ObjExcel = CreateObject("Excel.Application")
    '''   Set objLibro = ObjExcel.Workbooks.Add
    '''   Set objHoja = objLibro.ActiveSheet
    '''   '''ObjExcel.Visible = True
    '''   ObjExcel.Visible = False
    '''   objLibro.Sheets(1).Select
    '''   Set objHoja = objLibro.ActiveSheet
    '''   objLibro.ActiveSheet.Name = "Listado Ventas"
    '''
    '''   Encabezado
    '''   LlenaDatos
    '''
    '''   objLibro.SaveAs gstrCorpoDriveLocal & "\Sistema\Informes\" & Archivo & "", _
    ''''   FileFormat:=xlNormal, Password:="", writerespassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    '''   Screen.MousePointer = ccDefault
    '''   DoEvents
    '''   Select Case MsgBox("Se ha creado el archivo " & Archivo & " ¿Desea abrirlo?", vbYesNoCancel + vbQuestion, gstrNombCortoEmpresa)
    '''       Case vbYes:
    '''            ObjExcel.Visible = True
    '''            Set ObjExcel = Nothing
    '''            Set objLibro = Nothing
    '''            Set objHoja = Nothing
    '''       Case vbNo Or vbCancel:
    '''            CierraInstanciasdeExcel 1
    '''   End Select
    '''
    '''Err:
    '''   If Err.Number = 70 Then
    '''      MsgBox "No se puede generar un nuevo archivo hasta que el anterior este cerrado.", vbCritical + vbOKOnly, gstrNombCortoEmpresa
    '''      CierraInstanciasdeExcel 2
    '''      Screen.MousePointer = vbDefault
    '''   ElseIf Err.Number <> 0 Then
    '''      ModEstandar.MostrarError
    '''      CierraInstanciasdeExcel 1
    '''      Screen.MousePointer = vbDefault
    '''   End If
    '''End Sub

    '''11NOV2010 - MAVF
    '''06DIC2010 - MAVF
    Sub Encabezado()
        On Error GoTo Err_Renamed
        Dim Columna As Integer

        With objHoja
            Columna = 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Cliente"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 10

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Nombre"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 32

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Domicilio"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 28.71

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Colonia"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 23.57

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Ciudad"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 17.71

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "CP"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 8


            ''' 11NOV2010 - MAVF - COLUMNAS NUEVAS ***********************************************************
            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Tel Casa"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 17

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Tel Oficina"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 17

            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Email"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 8
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 35
            ''' *********************************************************************

            If chkMostrarImporte.CheckState Then
                Columna = Columna + 1
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "Impte Vtas"
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
                With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                    .Bold = True
                    .Size = 8
                    .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                    .Name = "Arial"
                End With
                .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 11.29
            End If


            '''11NOV2010 - MAVF - COLUMNA PARA INDICAR SI LE MISMO CLIENTE DE VENTAS TUVO REPARACIONES
            Columna = Columna + 1
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Select()
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna))._Default = "*"
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).VerticalAlignment = Excel.Constants.xlBottom
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).HorizontalAlignment = Excel.Constants.xlCenter
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).WrapText = False
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Interior.ColorIndex = 10
            With .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).Font
                .Bold = True
                .Size = 10
                .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
                .Name = "Arial"
            End With
            .Range(.Cells._Default(C_ENCABEZADO, Columna), .Cells._Default(C_ENCABEZADO, Columna)).ColumnWidth = 1

            '''FORMATO TIPO TEXTO PARA COLUMNAS CP-TEL1-TEL2
            '''06DIC2010 - MAVF
            .Range("F:F").Select()
            .Range("F:F").NumberFormat = "@"
            .Range("G:G").Select()
            .Range("G:G").NumberFormat = "@"
            .Range("H:H").Select()
            .Range("H:H").NumberFormat = "@"

        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    '''Sub Encabezado()
    '''On Error GoTo Err
    '''   Dim Columna    As Integer
    '''
    '''   With objHoja
    '''      .Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)).Select
    '''      .Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)) = "Cliente"
    '''      .Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)).VerticalAlignment = xlBottom
    '''      .Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)).HorizontalAlignment = xlCenter
    '''      .Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)).WrapText = False
    '''      .Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)).Interior.ColorIndex = 10
    '''      .Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''      With .Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)).Font
    '''          .Bold = True
    '''          .Size = 8
    '''          .Color = vbYellow
    '''          .Name = "Arial"
    '''      End With
    '''      .Range(.Cells(C_ENCABEZADO, 1), .Cells(C_ENCABEZADO, 1)).ColumnWidth = 10
    '''
    '''      .Range(.Cells(C_ENCABEZADO, 2), .Cells(C_ENCABEZADO, 2)).Select
    '''      .Range(.Cells(C_ENCABEZADO, 2), .Cells(C_ENCABEZADO, 2)) = "Nombre"
    '''      .Range(.Cells(C_ENCABEZADO, 2), .Cells(C_ENCABEZADO, 2)).VerticalAlignment = xlBottom
    '''      .Range(.Cells(C_ENCABEZADO, 2), .Cells(C_ENCABEZADO, 2)).HorizontalAlignment = xlCenter
    '''      .Range(.Cells(C_ENCABEZADO, 2), .Cells(C_ENCABEZADO, 2)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''      .Range(.Cells(C_ENCABEZADO, 2), .Cells(C_ENCABEZADO, 2)).WrapText = False
    '''      .Range(.Cells(C_ENCABEZADO, 2), .Cells(C_ENCABEZADO, 2)).Interior.ColorIndex = 10
    '''      With .Range(.Cells(C_ENCABEZADO, 2), .Cells(C_ENCABEZADO, 2)).Font
    '''         .Bold = True
    '''         .Size = 8
    '''         .Color = vbYellow
    '''         .Name = "Arial"
    '''      End With
    '''      .Range(.Cells(C_ENCABEZADO, 2), .Cells(C_ENCABEZADO, 2)).ColumnWidth = 32
    '''
    '''      .Range(.Cells(C_ENCABEZADO, 3), .Cells(C_ENCABEZADO, 3)).Select
    '''      .Range(.Cells(C_ENCABEZADO, 3), .Cells(C_ENCABEZADO, 3)) = "Domicilio"
    '''      .Range(.Cells(C_ENCABEZADO, 3), .Cells(C_ENCABEZADO, 3)).VerticalAlignment = xlBottom
    '''      .Range(.Cells(C_ENCABEZADO, 3), .Cells(C_ENCABEZADO, 3)).HorizontalAlignment = xlCenter
    '''      .Range(.Cells(C_ENCABEZADO, 3), .Cells(C_ENCABEZADO, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''      .Range(.Cells(C_ENCABEZADO, 3), .Cells(C_ENCABEZADO, 3)).WrapText = False
    '''      .Range(.Cells(C_ENCABEZADO, 3), .Cells(C_ENCABEZADO, 3)).Interior.ColorIndex = 10
    '''      With .Range(.Cells(C_ENCABEZADO, 3), .Cells(C_ENCABEZADO, 3)).Font
    '''         .Bold = True
    '''         .Size = 8
    '''         .Color = vbYellow
    '''         .Name = "Arial"
    '''      End With
    '''      .Range(.Cells(C_ENCABEZADO, 3), .Cells(C_ENCABEZADO, 3)).ColumnWidth = 28.71
    '''
    '''      .Range(.Cells(C_ENCABEZADO, 4), .Cells(C_ENCABEZADO, 4)).Select
    '''      .Range(.Cells(C_ENCABEZADO, 4), .Cells(C_ENCABEZADO, 4)) = "Colonia"
    '''      .Range(.Cells(C_ENCABEZADO, 4), .Cells(C_ENCABEZADO, 4)).VerticalAlignment = xlBottom
    '''      .Range(.Cells(C_ENCABEZADO, 4), .Cells(C_ENCABEZADO, 4)).HorizontalAlignment = xlCenter
    '''      .Range(.Cells(C_ENCABEZADO, 4), .Cells(C_ENCABEZADO, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''      .Range(.Cells(C_ENCABEZADO, 4), .Cells(C_ENCABEZADO, 4)).WrapText = False
    '''      .Range(.Cells(C_ENCABEZADO, 4), .Cells(C_ENCABEZADO, 4)).Interior.ColorIndex = 10
    '''      With .Range(.Cells(C_ENCABEZADO, 4), .Cells(C_ENCABEZADO, 4)).Font
    '''         .Bold = True
    '''         .Size = 8
    '''         .Color = vbYellow
    '''         .Name = "Arial"
    '''      End With
    '''      .Range(.Cells(C_ENCABEZADO, 4), .Cells(C_ENCABEZADO, 4)).ColumnWidth = 23.57
    '''
    '''      .Range(.Cells(C_ENCABEZADO, 5), .Cells(C_ENCABEZADO, 5)).Select
    '''      .Range(.Cells(C_ENCABEZADO, 5), .Cells(C_ENCABEZADO, 5)) = "Ciudad"
    '''      .Range(.Cells(C_ENCABEZADO, 5), .Cells(C_ENCABEZADO, 5)).VerticalAlignment = xlBottom
    '''      .Range(.Cells(C_ENCABEZADO, 5), .Cells(C_ENCABEZADO, 5)).HorizontalAlignment = xlCenter
    '''      .Range(.Cells(C_ENCABEZADO, 5), .Cells(C_ENCABEZADO, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''      .Range(.Cells(C_ENCABEZADO, 5), .Cells(C_ENCABEZADO, 5)).WrapText = False
    '''      .Range(.Cells(C_ENCABEZADO, 5), .Cells(C_ENCABEZADO, 5)).Interior.ColorIndex = 10
    '''      With .Range(.Cells(C_ENCABEZADO, 5), .Cells(C_ENCABEZADO, 5)).Font
    '''         .Bold = True
    '''         .Size = 8
    '''         .Color = vbYellow
    '''         .Name = "Arial"
    '''      End With
    '''      .Range(.Cells(C_ENCABEZADO, 5), .Cells(C_ENCABEZADO, 5)).ColumnWidth = 17.71
    '''
    '''      .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)).Select
    '''      .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)) = "CP"
    '''      .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)).VerticalAlignment = xlBottom
    '''      .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)).HorizontalAlignment = xlCenter
    '''      .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''      .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
    '''      .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)).WrapText = False
    '''      .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)).Interior.ColorIndex = 10
    '''      With .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)).Font
    '''         .Bold = True
    '''         .Size = 8
    '''         .Color = vbYellow
    '''         .Name = "Arial"
    '''      End With
    '''      .Range(.Cells(C_ENCABEZADO, 6), .Cells(C_ENCABEZADO, 6)).ColumnWidth = 8
    '''
    '''      If chkMostrarImporte.Value Then
    '''         .Range(.Cells(C_ENCABEZADO, 7), .Cells(C_ENCABEZADO, 7)).Select
    '''         .Range(.Cells(C_ENCABEZADO, 7), .Cells(C_ENCABEZADO, 7)) = "Impte Vtas"
    '''         .Range(.Cells(C_ENCABEZADO, 7), .Cells(C_ENCABEZADO, 7)).VerticalAlignment = xlBottom
    '''         .Range(.Cells(C_ENCABEZADO, 7), .Cells(C_ENCABEZADO, 7)).HorizontalAlignment = xlCenter
    '''         .Range(.Cells(C_ENCABEZADO, 7), .Cells(C_ENCABEZADO, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''         .Range(.Cells(C_ENCABEZADO, 7), .Cells(C_ENCABEZADO, 7)).WrapText = False
    '''         .Range(.Cells(C_ENCABEZADO, 7), .Cells(C_ENCABEZADO, 7)).Interior.ColorIndex = 10
    '''         With .Range(.Cells(C_ENCABEZADO, 7), .Cells(C_ENCABEZADO, 7)).Font
    '''            .Bold = True
    '''            .Size = 8
    '''            .Color = vbYellow
    '''            .Name = "Arial"
    '''         End With
    '''         .Range(.Cells(C_ENCABEZADO, 7), .Cells(C_ENCABEZADO, 7)).ColumnWidth = 11.29
    '''      End If
    '''   End With
    '''
    '''Err:
    '''   If Err.Number <> 0 Then
    '''      ModEstandar.MostrarError
    '''      CierraInstanciasdeExcel 1
    '''      Screen.MousePointer = vbDefault
    '''   End If
    '''End Sub

    Sub LlenaDatos()
        On Error GoTo Err_Renamed
        Dim Renglon As Integer
        Dim Columna As Integer
        Dim TotalReg As Integer

        With objHoja
            RsGral.MoveFirst()
            Columna = 1
            Renglon = 1
            TotalReg = RsGral.RecordCount
            Do While Not RsGral.EOF
                '''DETALLE
                Renglon = Renglon + 1
                Columna = 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("CodCliente").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Trim(RsGral.Fields("DescCliente").Value)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("Domicilio").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("Colonia").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("Ciudad").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("CP").Value
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                ''' 11NOV2010 - MAVF - COLUMNAS NUEVAS DEL REPORTE (TELCASA-TELOFICINA-EMAIL)
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Trim(RsGral.Fields("TelCasa").Value)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Trim(RsGral.Fields("TelOficina").Value)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With

                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = Trim(RsGral.Fields("Email").Value)
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Size = 8
                    .Name = "Arial"
                End With
                ''' ******************************************************************


                If chkMostrarImporte.CheckState Then
                    Columna = Columna + 1
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = RsGral.Fields("Impte").Value
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).NumberFormat = "###,##0.00"
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlRight
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                    If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                        .Size = 8
                        .Name = "Arial"
                    End With
                End If


                '''11NOV2010 - MAVF
                Columna = Columna + 1
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna))._Default = IIf(RsGral.Fields("Impte").Value = 0, "", IIf(RsGral.Fields("TipoC").Value > 0, "*", ""))
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Select()
                .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).HorizontalAlignment = Excel.Constants.xlLeft
                If (Renglon - 1) = TotalReg Then .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                With .Range(.Cells._Default(Renglon, Columna), .Cells._Default(Renglon, Columna)).Font
                    .Bold = True
                    .Size = 10
                    .Name = "Arial"
                    .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
                End With

                RsGral.MoveNext()
            Loop
            .Application.ActiveWindow.Zoom = 100
            .Range("A1").Select()

        End With

Err_Renamed:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            CierraInstanciasdeExcel(1)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    '''Sub LlenaDatos()
    '''On Error GoTo Err
    '''   Dim Renglon          As Integer
    '''   Dim Columna          As Integer
    '''   Dim TotalReg         As Long
    '''
    '''   With objHoja
    '''      RsGral.MoveFirst
    '''      Columna = 1
    '''      Renglon = 1
    '''      TotalReg = RsGral.RecordCount
    '''      Do While Not RsGral.EOF
    '''         '''DETALLE
    '''         Renglon = Renglon + 1
    '''         Columna = 1
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = RsGral!CodCliente
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Select
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlRight
    '''         If (Renglon - 1) = TotalReg Then .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''         With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
    '''            .Size = 8
    '''            .Name = "Arial"
    '''         End With
    '''
    '''         Columna = Columna + 1
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = Trim(RsGral!Nombre)
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Select
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlLeft
    '''         If (Renglon - 1) = TotalReg Then .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''         With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
    '''            .Size = 8
    '''            .Name = "Arial"
    '''         End With
    '''
    '''         Columna = Columna + 1
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = RsGral!Domicilio
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Select
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlLeft
    '''         If (Renglon - 1) = TotalReg Then .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''         With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
    '''            .Size = 8
    '''            .Name = "Arial"
    '''         End With
    '''
    '''         Columna = Columna + 1
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = RsGral!Colonia
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Select
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlLeft
    '''         If (Renglon - 1) = TotalReg Then .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''         With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
    '''            .Size = 8
    '''            .Name = "Arial"
    '''         End With
    '''
    '''         Columna = Columna + 1
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = RsGral!Ciudad
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Select
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlLeft
    '''         If (Renglon - 1) = TotalReg Then .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''         With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
    '''            .Size = 8
    '''            .Name = "Arial"
    '''         End With
    '''
    '''         Columna = Columna + 1
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = RsGral!CP
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Select
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlLeft
    '''         If (Renglon - 1) = TotalReg Then .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''         .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeRight).LineStyle = xlContinuous
    '''         With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
    '''            .Size = 8
    '''            .Name = "Arial"
    '''         End With
    '''
    '''         If chkMostrarImporte.Value Then
    '''            Columna = Columna + 1
    '''            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)) = RsGral!importe
    '''            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Select
    '''            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).NumberFormat = "###,##0.00"
    '''            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).HorizontalAlignment = xlRight
    '''            .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    '''            If (Renglon - 1) = TotalReg Then .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    '''            With .Range(.Cells(Renglon, Columna), .Cells(Renglon, Columna)).Font
    '''               .Size = 8
    '''               .Name = "Arial"
    '''            End With
    '''         End If
    '''         RsGral.MoveNext
    '''      Loop
    '''      .Application.ActiveWindow.Zoom = 100
    '''      .Range("A1").Select
    '''
    '''   End With
    '''
    '''Err:
    '''   If Err.Number <> 0 Then
    '''      ModEstandar.MostrarError
    '''      CierraInstanciasdeExcel 1
    '''      Screen.MousePointer = vbDefault
    '''   End If
    '''End Sub

    Sub CierraInstanciasdeExcel(ByRef Tipo As Integer)
        If Tipo = 1 Then
            objLibro.Close()
            ObjExcel.Quit()
        End If
        If ObjExcel Is Nothing Then ObjExcel = Nothing
        If objLibro Is Nothing Then objLibro = Nothing
        If objHoja Is Nothing Then objHoja = Nothing
    End Sub
    ''' 06OCT2006 - MAVF
    ''' FUNCIONES PARA GENERACION DE ARCHIVO A EXCEL **********************************************************************
    ''' *******************************************************************************************************************
    Private Sub chkTodas_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkTodas.CheckStateChanged
        Select Case Me.chkTodas.CheckState
            Case System.Windows.Forms.CheckState.Checked
                mblnFueraChange = True
                Me.dbcSucursal.Text = "[ Todas ... ]"
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

    '     SE OCULTO EL COMBO DEL CLIENTE - TAMBIEN SU CODIGO
    '     06OCT2006 - MAVF
    '    Private Sub dbcCliente_Change()
    '        On Local Error GoTo Merr
    '        Dim lStrSql As String

    '        If mblnFueraChange Then Exit Sub

    '        lStrSql = "SELECT CodCliente, LTrim(RTrim(descCliente)) as descCliente FROM CatClientes Where DescCliente LIKE '" & Trim(Me.dbcCliente.Text) & "%'"
    '        ModDCombo.DCChange lStrSql, tecla, Me.dbcCliente

    '        If Trim(Me.dbcCliente.Text) = "" Then
    '            mintCodCliente = 0
    '            dbcCliente_LostFocus()
    '        End If

    'Merr:
    '        If Err.Number <> 0 Then
    '            ModEstandar.MostrarError()
    '        End If
    '    End Sub

    '    Private Sub dbcCliente_GotFocus()
    '        Pon_Tool()
    '        gStrSql = "SELECT CodCliente, LTrim(RTrim(DescCliente)) as DescCliente FROM CatClientes "
    '        ModDCombo.DCGotFocus gStrSql, Me.dbcCliente
    '    End Sub

    '    Private Sub dbcCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    '        If KeyCode = vbKeyEscape Then
    '            If Me.dbcSucursal.Enabled Then
    '                Me.dbcSucursal.SetFocus
    '            Else
    '                Me.chkTodas.SetFocus
    '            End If
    '            KeyCode = 0
    '        End If
    '        tecla = KeyCode
    '    End Sub

    '    Private Sub dbcCliente_KeyUp(KeyCode As Integer, Shift As Integer)
    '        '    Dim Aux As String
    '        '    Aux = Trim(Me.dbcCliente.text)
    '        '    If Me.dbcCliente.SelectedItem <> 0 Then
    '        '        dbcCliente_LostFocus
    '        '    End If
    '        '    Me.dbcCliente.text = Aux
    '    End Sub

    '    Private Sub dbcCliente_LostFocus()
    '        Dim Aux As Long
    '        If Screen.ActiveForm.Name <> Me.Name Then
    '            Exit Sub
    '        End If
    '        gStrSql = "SELECT CodCliente, LTrim(RTrim(DescCliente)) as DescCliente FROM CatClientes Where DescCliente LIKE '" & Trim(Me.dbcCliente.Text) & "%'"
    '        Aux = mintCodCliente
    '        mintCodCliente = 0
    '        If Trim(Me.dbcCliente.Text) <> Trim(C_TODOS) Or Trim(Me.dbcCliente.Text) = "" Then
    '            ModDCombo.DCLostFocus Me.dbcCliente, gStrSql, mintCodCliente
    '        End If

    '        If Aux <> mintCodCliente Then
    '            If mintCodCliente = 0 Then
    '                mblnFueraChange = True
    '                Me.dbcCliente.Text = C_TODOS
    '                Me.dbcCliente.Enabled = True
    '                mblnFueraChange = False
    '            End If
    '        End If
    '        If Trim(Me.dbcCliente.Text) = "" Then Me.dbcCliente.Text = C_TODOS
    '    End Sub

    '    Private Sub dbcCliente_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '        '    Dim Aux As String
    '        '    Aux = Trim(Me.dbcCliente.text)
    '        '    If Me.dbcCliente.SelectedItem <> 0 Then
    '        '        dbcCliente_LostFocus
    '        '    End If
    '        '    Me.dbcCliente.text = Aux
    '    End Sub

    Private Sub dbcSucursal_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.CursorChanged
        On Error GoTo Merr
        Dim lStrSql As String

        If mblnFueraChange Then Exit Sub

        lStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        ModDCombo.DCChange(lStrSql, tecla, dbcSucursal)

        If Trim(Me.dbcSucursal.Text) = "" Then
            mintCodSucursal = 0
            'dbcSucursal_LostFocus
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

    Private Sub dbcSucursal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursal.Leave
        Dim I As Integer
        Dim Aux As Integer
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
        '    Exit Sub
        'Else
        '    If Trim(Me.dbcSucursal.Text) = "" Or Trim(Me.dbcSucursal.Text) = C_TODAS Then Exit Sub
        'End If
        gStrSql = "SELECT codAlmacen, LTrim(RTrim(descAlmacen)) as descAlmacen FROM catAlmacen Where TipoAlmacen = 'P' and descAlmacen LIKE '" & Trim(Me.dbcSucursal.Text) & "%'"
        Aux = mintCodSucursal
        mintCodSucursal = 0
        ModDCombo.DCLostFocus((Me.dbcSucursal), gStrSql, mintCodSucursal)
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
        'msglTiempoCambioF = VB.Timer()
    End Sub

    Private Sub frmVtasRPTVtasSalMciaListadoVtasxCte_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmVtasRPTVtasSalMciaListadoVtasxCte_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmVtasRPTVtasSalMciaListadoVtasxCte_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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

    Private Sub frmVtasRPTVtasSalMciaListadoVtasxCte_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma el apóstrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte letras en mayúsculas
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmVtasRPTVtasSalMciaListadoVtasxCte_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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

    Private Sub frmVtasRPTVtasSalMciaListadoVtasxCte_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub frmVtasRPTVtasSalMciaListadoVtasxCte_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'cmd.CommandTimeout = 90
        'Me = Nothing
        IsNothing(Me)
    End Sub

    Private Sub txtMensaje_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMensaje.Enter
        Pon_Tool()
        ModEstandar.SelTxt()
    End Sub

    Private Sub btnNuevo_Click(sender As Object, e As EventArgs) Handles btnNuevo.Click
        Nuevo()
    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Imprime()
    End Sub

    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtMensaje = New System.Windows.Forms.TextBox()
        Me.chkMostrarImporte = New System.Windows.Forms.CheckBox()
        Me._fraVtas_0 = New System.Windows.Forms.GroupBox()
        Me.chkTodas = New System.Windows.Forms.CheckBox()
        Me.dbcSucursal = New System.Windows.Forms.ComboBox()
        Me._lblVentas_0 = New System.Windows.Forms.Label()
        Me._fraVtas_1 = New System.Windows.Forms.GroupBox()
        Me.dtpDesde = New System.Windows.Forms.DateTimePicker()
        Me.dtpHasta = New System.Windows.Forms.DateTimePicker()
        Me._lblVentas_1 = New System.Windows.Forms.Label()
        Me._lblVentas_2 = New System.Windows.Forms.Label()
        Me.chkImpuesto = New System.Windows.Forms.CheckBox()
        Me.dbcCliente = New System.Windows.Forms.ComboBox()
        Me._lblVentas_5 = New System.Windows.Forms.Label()
        Me._lblRpt_2 = New System.Windows.Forms.Label()
        Me.fraVtas = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.lblRpt = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblVentas = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.btnNuevo = New System.Windows.Forms.Button()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me._fraVtas_0.SuspendLayout()
        Me._fraVtas_1.SuspendLayout()
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
        Me.txtMensaje.Location = New System.Drawing.Point(12, 227)
        Me.txtMensaje.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMensaje.MaxLength = 100
        Me.txtMensaje.Multiline = True
        Me.txtMensaje.Name = "txtMensaje"
        Me.txtMensaje.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMensaje.Size = New System.Drawing.Size(335, 69)
        Me.txtMensaje.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.txtMensaje, "Mensaje que aparecerá en el encabezado del  reporte")
        '
        'chkMostrarImporte
        '
        Me.chkMostrarImporte.BackColor = System.Drawing.SystemColors.Control
        Me.chkMostrarImporte.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMostrarImporte.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMostrarImporte.Location = New System.Drawing.Point(11, 181)
        Me.chkMostrarImporte.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMostrarImporte.Name = "chkMostrarImporte"
        Me.chkMostrarImporte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMostrarImporte.Size = New System.Drawing.Size(120, 18)
        Me.chkMostrarImporte.TabIndex = 12
        Me.chkMostrarImporte.Text = "Mostrar Importe"
        Me.chkMostrarImporte.UseVisualStyleBackColor = False
        '
        '_fraVtas_0
        '
        Me._fraVtas_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_0.Controls.Add(Me.chkTodas)
        Me._fraVtas_0.Controls.Add(Me.dbcSucursal)
        Me._fraVtas_0.Controls.Add(Me._lblVentas_0)
        Me._fraVtas_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_0.Location = New System.Drawing.Point(6, 6)
        Me._fraVtas_0.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.Name = "_fraVtas_0"
        Me._fraVtas_0.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_0.Size = New System.Drawing.Size(315, 47)
        Me._fraVtas_0.TabIndex = 0
        Me._fraVtas_0.TabStop = False
        '
        'chkTodas
        '
        Me.chkTodas.BackColor = System.Drawing.SystemColors.Control
        Me.chkTodas.Checked = True
        Me.chkTodas.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTodas.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTodas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkTodas.Location = New System.Drawing.Point(6, 0)
        Me.chkTodas.Margin = New System.Windows.Forms.Padding(2)
        Me.chkTodas.Name = "chkTodas"
        Me.chkTodas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTodas.Size = New System.Drawing.Size(161, 17)
        Me.chkTodas.TabIndex = 1
        Me.chkTodas.Text = "Todas las sucursales"
        Me.chkTodas.UseVisualStyleBackColor = False
        '
        'dbcSucursal
        '
        Me.dbcSucursal.Location = New System.Drawing.Point(75, 17)
        Me.dbcSucursal.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcSucursal.Name = "dbcSucursal"
        Me.dbcSucursal.Size = New System.Drawing.Size(188, 21)
        Me.dbcSucursal.TabIndex = 3
        '
        '_lblVentas_0
        '
        Me._lblVentas_0.AutoSize = True
        Me._lblVentas_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_0.Location = New System.Drawing.Point(20, 20)
        Me._lblVentas_0.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_0.Name = "_lblVentas_0"
        Me._lblVentas_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_0.Size = New System.Drawing.Size(51, 13)
        Me._lblVentas_0.TabIndex = 2
        Me._lblVentas_0.Text = "Sucursal:"
        '
        '_fraVtas_1
        '
        Me._fraVtas_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraVtas_1.Controls.Add(Me.dtpDesde)
        Me._fraVtas_1.Controls.Add(Me.dtpHasta)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_1)
        Me._fraVtas_1.Controls.Add(Me._lblVentas_2)
        Me._fraVtas_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._fraVtas_1.Location = New System.Drawing.Point(11, 92)
        Me._fraVtas_1.Margin = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.Name = "_fraVtas_1"
        Me._fraVtas_1.Padding = New System.Windows.Forms.Padding(2)
        Me._fraVtas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraVtas_1.Size = New System.Drawing.Size(335, 50)
        Me._fraVtas_1.TabIndex = 6
        Me._fraVtas_1.TabStop = False
        Me._fraVtas_1.Text = "Período ..."
        '
        'dtpDesde
        '
        Me.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDesde.Location = New System.Drawing.Point(62, 19)
        Me.dtpDesde.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpDesde.Name = "dtpDesde"
        Me.dtpDesde.Size = New System.Drawing.Size(100, 20)
        Me.dtpDesde.TabIndex = 8
        '
        'dtpHasta
        '
        Me.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpHasta.Location = New System.Drawing.Point(226, 18)
        Me.dtpHasta.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpHasta.Name = "dtpHasta"
        Me.dtpHasta.Size = New System.Drawing.Size(97, 20)
        Me.dtpHasta.TabIndex = 10
        '
        '_lblVentas_1
        '
        Me._lblVentas_1.AutoSize = True
        Me._lblVentas_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_1.Location = New System.Drawing.Point(12, 24)
        Me._lblVentas_1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_1.Name = "_lblVentas_1"
        Me._lblVentas_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_1.Size = New System.Drawing.Size(52, 13)
        Me._lblVentas_1.TabIndex = 7
        Me._lblVentas_1.Text = "Desde el "
        '
        '_lblVentas_2
        '
        Me._lblVentas_2.AutoSize = True
        Me._lblVentas_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_2.Location = New System.Drawing.Point(176, 23)
        Me._lblVentas_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_2.Name = "_lblVentas_2"
        Me._lblVentas_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_2.Size = New System.Drawing.Size(46, 13)
        Me._lblVentas_2.TabIndex = 9
        Me._lblVentas_2.Text = "Hasta el"
        '
        'chkImpuesto
        '
        Me.chkImpuesto.BackColor = System.Drawing.SystemColors.Control
        Me.chkImpuesto.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkImpuesto.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkImpuesto.Location = New System.Drawing.Point(11, 156)
        Me.chkImpuesto.Margin = New System.Windows.Forms.Padding(2)
        Me.chkImpuesto.Name = "chkImpuesto"
        Me.chkImpuesto.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkImpuesto.Size = New System.Drawing.Size(136, 20)
        Me.chkImpuesto.TabIndex = 11
        Me.chkImpuesto.Text = "Incluir Impuesto"
        Me.chkImpuesto.UseVisualStyleBackColor = False
        '
        'dbcCliente
        '
        Me.dbcCliente.Location = New System.Drawing.Point(52, 57)
        Me.dbcCliente.Margin = New System.Windows.Forms.Padding(2)
        Me.dbcCliente.Name = "dbcCliente"
        Me.dbcCliente.Size = New System.Drawing.Size(217, 21)
        Me.dbcCliente.TabIndex = 5
        Me.dbcCliente.Visible = False
        '
        '_lblVentas_5
        '
        Me._lblVentas_5.AutoSize = True
        Me._lblVentas_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblVentas_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblVentas_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblVentas_5.Location = New System.Drawing.Point(10, 61)
        Me._lblVentas_5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblVentas_5.Name = "_lblVentas_5"
        Me._lblVentas_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblVentas_5.Size = New System.Drawing.Size(39, 13)
        Me._lblVentas_5.TabIndex = 4
        Me._lblVentas_5.Text = "Cliente"
        Me._lblVentas_5.Visible = False
        '
        '_lblRpt_2
        '
        Me._lblRpt_2.AutoSize = True
        Me._lblRpt_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblRpt_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblRpt_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblRpt_2.Location = New System.Drawing.Point(12, 203)
        Me._lblRpt_2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me._lblRpt_2.Name = "_lblRpt_2"
        Me._lblRpt_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblRpt_2.Size = New System.Drawing.Size(175, 13)
        Me._lblRpt_2.TabIndex = 13
        Me._lblRpt_2.Text = "Mensaje adicional para el reporte ..."
        '
        'btnNuevo
        '
        Me.btnNuevo.BackColor = System.Drawing.SystemColors.Control
        Me.btnNuevo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNuevo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNuevo.Location = New System.Drawing.Point(127, 311)
        Me.btnNuevo.Name = "btnNuevo"
        Me.btnNuevo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNuevo.Size = New System.Drawing.Size(109, 36)
        Me.btnNuevo.TabIndex = 79
        Me.btnNuevo.Text = "&Nuevo"
        Me.btnNuevo.UseVisualStyleBackColor = False
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.SystemColors.Control
        Me.btnImprimir.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnImprimir.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnImprimir.Location = New System.Drawing.Point(12, 311)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnImprimir.Size = New System.Drawing.Size(109, 36)
        Me.btnImprimir.TabIndex = 78
        Me.btnImprimir.Text = "&Imprimir"
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'frmVtasRPTVtasSalMciaListadoVtasxCte
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(359, 356)
        Me.Controls.Add(Me.btnNuevo)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.chkMostrarImporte)
        Me.Controls.Add(Me._fraVtas_0)
        Me.Controls.Add(Me._fraVtas_1)
        Me.Controls.Add(Me.chkImpuesto)
        Me.Controls.Add(Me.txtMensaje)
        Me.Controls.Add(Me.dbcCliente)
        Me.Controls.Add(Me._lblVentas_5)
        Me.Controls.Add(Me._lblRpt_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(260, 235)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "frmVtasRPTVtasSalMciaListadoVtasxCte"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Listado de Ventas - Etiquetas"
        Me._fraVtas_0.ResumeLayout(False)
        Me._fraVtas_0.PerformLayout()
        Me._fraVtas_1.ResumeLayout(False)
        Me._fraVtas_1.PerformLayout()
        CType(Me.fraVtas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRpt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

End Class