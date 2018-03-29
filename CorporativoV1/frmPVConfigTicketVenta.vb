Option Strict Off
Option Explicit On
Public Class frmPVConfigTicketVenta
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '**********************************************************************************************************************'
    '*PROGRAMA :             CONFIGURACION DEL TICKET DE VENTA                                                            *'
    '*AUTOR :                JUAN CARLOS OSUNA CORRALES                                                                   *'
    '*FECHA DE INICIO :      JUEVES 22 DE MAYO DE 2003                                                                    *'
    '*FECHA DE TERMINACION :                                                                                              *'
    '**********************************************************************************************************************'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim isLoad As Boolean = False
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkAplicarSucursales As System.Windows.Forms.CheckBox
    Public WithEvents optCredito As System.Windows.Forms.RadioButton
    Public WithEvents optContado As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents TxtSaltos As System.Windows.Forms.TextBox
    Public WithEvents TxtEtiqueta As System.Windows.Forms.TextBox
    Public WithEvents btnPrueba As System.Windows.Forms.Button
    Public WithEvents _FlexDetalle_0 As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents TxtCols As System.Windows.Forms.TextBox
    Public WithEvents _lblFormula_0 As System.Windows.Forms.Label
    Public WithEvents _Label1_1 As System.Windows.Forms.Label
    Public WithEvents _Marco_1 As System.Windows.Forms.GroupBox
    Public WithEvents _btnDatosEnc_14 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_16 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_15 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_13 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_12 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_11 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_10 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_9 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_8 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_7 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_6 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_5 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_4 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_3 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_1 As System.Windows.Forms.Button
    Public WithEvents _btnDatosEnc_0 As System.Windows.Forms.Button
    Public WithEvents _Marco_0 As System.Windows.Forms.GroupBox
    Public WithEvents TxtColumna As System.Windows.Forms.TextBox
    Public WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
    Public WithEvents _btnDatosDet_13 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_12 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_11 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_10 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_7 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_6 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_5 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_4 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_3 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_2 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_1 As System.Windows.Forms.Button
    Public WithEvents _btnDatosDet_0 As System.Windows.Forms.Button
    Public WithEvents _Marco_3 As System.Windows.Forms.GroupBox
    Public WithEvents _FlexDetalle_1 As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents _lblFormula_1 As System.Windows.Forms.Label
    Public WithEvents _Marco_2 As System.Windows.Forms.GroupBox
    Public WithEvents _SSTab1_TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents _btnDatosTot_12 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_16 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_14 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_15 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_13 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_11 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_10 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_9 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_8 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_7 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_6 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_5 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_4 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_3 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_2 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_35 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_1 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_0 As System.Windows.Forms.Button
    Public WithEvents fraTotalesContado As System.Windows.Forms.GroupBox
    Public WithEvents _btnDatosTot_30 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_29 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_34 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_32 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_33 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_31 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_28 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_27 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_26 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_25 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_24 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_23 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_22 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_21 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_20 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_19 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_18 As System.Windows.Forms.Button
    Public WithEvents _btnDatosTot_17 As System.Windows.Forms.Button
    Public WithEvents fraTotalesCredito As System.Windows.Forms.GroupBox
    Public WithEvents _FlexDetalle_2 As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Public WithEvents _lblFormula_2 As System.Windows.Forms.Label
    Public WithEvents _Marco_4 As System.Windows.Forms.GroupBox
    Public WithEvents _SSTab1_TabPage2 As System.Windows.Forms.TabPage
    Public WithEvents SSTab1 As System.Windows.Forms.TabControl
    Public WithEvents dbcSucursales As System.Windows.Forms.ComboBox
    Public WithEvents _Label1_0 As System.Windows.Forms.Label
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents _Label2_0 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents FlexDetalle As AxMSHFlexGridArray.AxMSHFlexGridArray
    Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents Marco As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents btnDatosDet As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Public WithEvents btnDatosEnc As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Public WithEvents btnDatosTot As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Public WithEvents lblFormula As Microsoft.VisualBasic.Compatibility.VB6.LabelArray


    Dim mblnSALIR As Boolean
    Dim mblnNuevo As Boolean
    Dim mblnSaliryGrabar As Boolean
    Dim mblnPierdeFoco As Boolean
    Dim mstrGrupo As String 'Contiene una letra que indica el Grupo en el que se está grabando, la configuracion (Encabezado "E", Detalle "D", Totales "T")
    Dim FueraChange As Boolean
    Dim tecla As Integer
    Dim intCodSucursal As Integer
    Dim auxCodSucursal As Integer

    Function Guardar() As Boolean
        'On Error GoTo Merr
        Dim blnTransaccion As Boolean
        Dim strRenglon As String
        Dim strEtiqueta As String
        Dim strFormula As String
        Dim strColumna As String
        Dim strSaltos As String
        Dim strGrupo As String
        Dim strTipo As String 'Credito ("CR") o Contado ("CO")
        Dim J, I, Ren As Object
        Dim Num As Integer
        Dim rsSucursales As ADODB.Recordset

        TxtSaltos_Leave(TxtSaltos, New System.EventArgs())
        TxtColumna_Leave(TxtColumna, New System.EventArgs())
        TxtEtiqueta_Leave(TxtEtiqueta, New System.EventArgs())
        Guardar = False
        If Cambios() = False And chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked Then
            mblnSaliryGrabar = True
            Me.Close()
            Exit Function
        End If
        'Valida si todos los datos han sido llenados para poder ser guardados
        If ValidaDatos() = False Then
            Exit Function
        End If
        'Verificar si se seleccionó el Check de Todas las Sucursales
        If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
            gStrSql = "SELECT CodAlmacen from catAlmacen Where TipoAlmacen = 'P' "
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
            rsSucursales = Cmd.Execute
        End If
        If optContado.Checked = True Then
            strTipo = "CO"
        ElseIf optCredito.Checked = True Then
            strTipo = "CR"
        End If
        Cnn.BeginTrans()
        blnTransaccion = True
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        If mblnNuevo Then
            strRenglon = CStr(1)
            strEtiqueta = ""
            strFormula = ""
            strColumna = TxtCols.Text
            strSaltos = CStr(0)
            strGrupo = "C"
            Err.Clear()
            ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(intCodSucursal), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_INSERCION, CStr(0))
            Cmd.Execute()
            If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
                For Num = 1 To rsSucursales.RecordCount
                    If rsSucursales.Fields("CodAlmacen").Value <> intCodSucursal Then
                        ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(rsSucursales.Fields("CodAlmacen").Value), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_INSERCION, CStr(0))
                        Cmd.Execute()
                    End If
                Next
            End If
            For J = 0 To 2
                With FlexDetalle(J)
                    'Obtener el GRupo en el que se está guardando la información
                    'UPGRADE_WARNING: Couldn't resolve default property of object J. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If J = 0 Then
                        strGrupo = "E"
                        'UPGRADE_WARNING: Couldn't resolve default property of object J. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ElseIf J = 1 Then
                        strGrupo = "D"
                        'UPGRADE_WARNING: Couldn't resolve default property of object J. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ElseIf J = 2 Then
                        strGrupo = "T"
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object Ren. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Ren = 1
                    For I = 1 To .Rows - 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If .get_TextMatrix(I, 0) <> "" Or .get_TextMatrix(I, 1) <> "" Or .get_TextMatrix(I, 2) <> "" Or .get_TextMatrix(I, 3) <> "" Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Ren. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strRenglon = Ren
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strEtiqueta = .get_TextMatrix(I, 0)
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strFormula = .get_TextMatrix(I, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strColumna = .get_TextMatrix(I, 2)
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strSaltos = .get_TextMatrix(I, 3)
                            ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(intCodSucursal), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_INSERCION, CStr(0))
                            Cmd.Execute()
                            If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
                                For Num = 1 To rsSucursales.RecordCount
                                    If rsSucursales.Fields("CodAlmacen").Value <> intCodSucursal Then
                                        ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(rsSucursales.Fields("CodAlmacen").Value), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_INSERCION, CStr(0))
                                        Cmd.Execute()
                                    End If
                                Next
                            End If
                            'UPGRADE_WARNING: Couldn't resolve default property of object Ren. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Ren = Ren + 1
                        End If
                    Next
                End With
            Next
        Else
            '        gStrSql = "Delete From ConfigTicketVenta Where Tipo= '" & strTipo & "' and CodAlmacen = " & intCodSucursal
            '        ModEstandar.BorraCmd
            '        Cmd.CommandText = "dbo.Up_Select_Datos"
            '        Cmd.CommandType = adCmdStoredProc
            '        Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
            '        Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
            ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(intCodSucursal), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_ELIMINACION, CStr(1))
            Cmd.Execute()
            If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
                For Num = 1 To rsSucursales.RecordCount
                    If rsSucursales.Fields("CodAlmacen").Value <> intCodSucursal Then
                        '                    gStrSql = "Delete From ConfigTicketVenta Where Tipo= '" & strTipo & "' and CodAlmacen = " & CInt(rsSucursales!CodAlmacen)
                        '                    ModEstandar.BorraCmd
                        '                    Cmd.CommandText = "dbo.Up_Select_Datos"
                        '                    Cmd.CommandType = adCmdStoredProc
                        '                    Cmd.Parameters.Append Cmd.CreateParameter("Renglon", adInteger, adParamReturnValue)
                        '                    Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql)
                        ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(CInt(rsSucursales.Fields("CodAlmacen").Value)), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_ELIMINACION, CStr(1))
                        Cmd.Execute()
                    End If
                    rsSucursales.MoveNext()
                Next
            End If

            strRenglon = CStr(1)
            strEtiqueta = ""
            strFormula = ""
            strColumna = TxtCols.Text
            strSaltos = CStr(0)
            strGrupo = "C"
            ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(intCodSucursal), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_INSERCION, CStr(0))
            Cmd.Execute()

            If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
                rsSucursales.MoveFirst()
                For Num = 1 To rsSucursales.RecordCount
                    If rsSucursales.Fields("CodAlmacen").Value <> intCodSucursal Then
                        ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(rsSucursales.Fields("CodAlmacen").Value), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_INSERCION, CStr(0))
                        Cmd.Execute()
                    End If
                    rsSucursales.MoveNext()
                Next
            End If
            For J = 0 To 2
                'Obtener el GRupo en el que se está guardando la información
                'UPGRADE_WARNING: Couldn't resolve default property of object J. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If J = 0 Then
                    strGrupo = "E"
                    'UPGRADE_WARNING: Couldn't resolve default property of object J. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf J = 1 Then
                    strGrupo = "D"
                    'UPGRADE_WARNING: Couldn't resolve default property of object J. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf J = 2 Then
                    strGrupo = "T"
                End If
                With FlexDetalle(J)
                    'UPGRADE_WARNING: Couldn't resolve default property of object Ren. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Ren = 1
                    For I = 1 To .Rows - 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If .get_TextMatrix(I, 0) <> "" Or .get_TextMatrix(I, 1) <> "" Or .get_TextMatrix(I, 2) <> "" Or .get_TextMatrix(I, 3) <> "" Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Ren. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strRenglon = Ren
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strEtiqueta = .get_TextMatrix(I, 0)
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strFormula = .get_TextMatrix(I, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strColumna = .get_TextMatrix(I, 2)
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strSaltos = .get_TextMatrix(I, 3)
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .set_TextMatrix(I, 4, .get_TextMatrix(I, 0))
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .set_TextMatrix(I, 5, .get_TextMatrix(I, 1))
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .set_TextMatrix(I, 6, .get_TextMatrix(I, 2))
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .set_TextMatrix(I, 7, .get_TextMatrix(I, 3))
                            ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(intCodSucursal), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_INSERCION, CStr(0))
                            Cmd.Execute()
                            If chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Checked Then
                                rsSucursales.MoveFirst()
                                For Num = 1 To rsSucursales.RecordCount
                                    If rsSucursales.Fields("CodAlmacen").Value <> intCodSucursal Then
                                        ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(rsSucursales.Fields("CodAlmacen").Value), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_INSERCION, CStr(0))
                                        Cmd.Execute()
                                    End If
                                    rsSucursales.MoveNext()
                                Next
                            End If

                            'UPGRADE_WARNING: Couldn't resolve default property of object Ren. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Ren = Ren + 1
                        End If
                    Next
                End With
            Next
        End If
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
        If mblnNuevo Then
            MsgBox("La Configuración del Ticket Ha sido Grabada Correctamente", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mensaje")
        Else
            MsgBox(C_msgACTUALIZADO, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ModVariables.gstrNombCortoEmpresa)
        End If
        '    InicializaVariables
        Guardar = True
        mblnSaliryGrabar = True
        '''Limpiar
        Exit Function
Merr:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Function

    Function Cambios() As Boolean
        Dim I As Object
        Dim J As Integer
        Cambios = True
        If Val(Trim(TxtCols.Text)) <> Val(TxtCols.Tag) Then Exit Function
        For J = 0 To 2
            With FlexDetalle(J)
                For I = 1 To .Rows - 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If .get_TextMatrix(I, 0) = "" And .get_TextMatrix(I, 1) = "" And .get_TextMatrix(I, 2) = "" And .get_TextMatrix(I, 3) = "" Then
                        Exit For
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If .get_TextMatrix(I, 0) <> .get_TextMatrix(I, 4) Then Exit Function
                        'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If .get_TextMatrix(I, 1) <> .get_TextMatrix(I, 5) Then Exit Function
                        'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If .get_TextMatrix(I, 2) <> .get_TextMatrix(I, 6) Then Exit Function
                        'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If .get_TextMatrix(I, 3) <> .get_TextMatrix(I, 7) Then Exit Function
                    End If
                Next
            End With
        Next
        Cambios = False
    End Function

    Sub Limpiar()
        '    InicializaVariables
        Nuevo()
        FueraChange = True
        dbcSucursales.Text = ""
        '''dbcSucursales.SetFocus
        FueraChange = False
    End Sub

    Sub Nuevo()
        If (Not isLoad) Then
            Exit Sub
        End If
        InicializaVariables()
        chkAplicarSucursales.CheckState = System.Windows.Forms.CheckState.Unchecked
        TxtCols.Text = ""
        TxtCols.Tag = ""
        FlexDetalle(0).Clear()
        FlexDetalle(1).Clear()
        FlexDetalle(2).Clear()
        Encabezado0(0)
        Encabezado1(1)
        Encabezado2(2)
    End Sub

    Function ValidaDatos() As Boolean
        Dim I As Object
        Dim J As Integer
        Dim Vacio As Boolean
        ValidaDatos = False
        If Val(TxtCols.Text) = 0 Then
            MsgBox("Proporcione el Numero de Columnas Para el Ticket", MsgBoxStyle.Information, gstrNombCortoEmpresa)
            TxtCols.Focus()
            Exit Function
        End If
        For J = 0 To 2
            With FlexDetalle(J)
                Vacio = True
                For I = 1 To .Rows - 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If (.get_TextMatrix(I, 0) <> "" Or .get_TextMatrix(I, 1) <> "") And (.get_TextMatrix(I, 2) <> "" And .get_TextMatrix(I, 3) <> "") Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Val(.get_TextMatrix(I, 2)) > Val(TxtCols.Text) Then
                            MsgBox("La Columna Especificada No Debe Excecer al Número de Cólumnas.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                            SSTab1.SelectedIndex = J
                            .Col = 2
                            'UPGRADE_WARNING: Couldn't resolve default property of object I. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .Row = I
                            .Focus()
                            Exit Function
                        End If
                        Vacio = False
                        'Exit For
                    Else
                        '                    MsgBox "Faltan Datos en la Pestaña " & SSTab1.TabCaption(j), vbOKOnly + vbInformation, gstrNombCortoEmpresa
                        '                    SSTab1.Tab = j
                        '                    .Row = i
                        '                    .SetFocus
                        '                    Exit Function
                    End If
                Next
                If Vacio And J = 0 Then
                    MsgBox("Proporcione la Configuración del Encabezado", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    SSTab1.SelectedIndex = J
                    FlexDetalle(J).Focus()
                    Exit Function
                ElseIf Vacio And J = 1 Then
                    MsgBox("Proporcione la Configuración del Detalle", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    SSTab1.SelectedIndex = J
                    FlexDetalle(J).Focus()
                    Exit Function
                ElseIf Vacio And J = 2 Then
                    MsgBox("Proporcione la Configuración de los Totales", MsgBoxStyle.Information, gstrNombCortoEmpresa)
                    SSTab1.SelectedIndex = J
                    FlexDetalle(J).Focus()
                    Exit Function
                End If
            End With
        Next
        ValidaDatos = True
    End Function

    Sub LlenaDatos()
        'On Error GoTo Merr
        Dim J As Object
        Dim Ren As Integer
        Dim Grupo As String
        Dim IndexStab As Integer
        '    Dim SeMostroCols As Boolean
        '    SeMostroCols = False
        With RsGral
            Do While Not .EOF
                Grupo = RsGral.Fields("Grupo").Value
                If Grupo = "C" Then
                    TxtCols.Text = RsGral.Fields("Columna").Value
                    TxtCols.Tag = RsGral.Fields("Columna").Value
                    RsGral.MoveNext()
                End If
                Grupo = RsGral.Fields("Grupo").Value
                Select Case Grupo
                    Case "E"
                        IndexStab = 0
                    Case "D"
                        IndexStab = 1
                    Case "T"
                        IndexStab = 2
                End Select
                Ren = RsGral.Fields("Renglon").Value
                With FlexDetalle(IndexStab)
                    If Ren >= .Rows - 1 Then
                        .Rows = Ren + 2
                    End If
                    .Col = 0
                    .set_TextMatrix(.Row, 0, RsGral.Fields("Etiqueta").Value)
                    .set_TextMatrix(.Row, 4, RsGral.Fields("Etiqueta").Value)
                    .Col = 1
                    .CellAlignment = 0
                    .set_TextMatrix(.Row, 1, RsGral.Fields("Formula").Value)
                    .set_TextMatrix(.Row, 5, RsGral.Fields("Formula").Value)
                    .set_TextMatrix(.Row, 2, RsGral.Fields("Columna").Value)
                    .set_TextMatrix(.Row, 6, RsGral.Fields("Columna").Value)
                    .set_TextMatrix(.Row, 3, RsGral.Fields("Saltos").Value)
                    .set_TextMatrix(.Row, 7, RsGral.Fields("Saltos").Value)
                    .set_TextMatrix(.Row, 8, RsGral.Fields("Renglon").Value)

                    .set_TextMatrix(.Row, 9, RsGral.Fields("Grupo").Value)
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Row = .Row + 1
                End With
                .MoveNext()
            Loop
        End With
        FlexDetalle(0).Col = 0
        FlexDetalle(0).Row = 1
        FlexDetalle(1).Col = 0
        FlexDetalle(1).Row = 1
        FlexDetalle(2).Col = 0
        FlexDetalle(2).Row = 1
        '    For J = 0 To 2
        '        With FlexDetalle(J)
        '            .Row = 1
        '            .Col = 0
        '        End With
        '    Next



        '    J = 0
        '    With RsGral
        '        TxtCols = RsGral!Columna
        '        TxtCols.Tag = RsGral!Columna
        '        RsGral.MoveNext
        '        Do While Not .EOF
        '            Ren = RsGral!Renglon
        '            With FlexDetalle(J)
        '                 If Ren > .Rows - 1 Then
        '                    .Rows = Ren + 1
        '                 End If
        '                 .Col = 0
        '                 .TextMatrix(.Row, 0) = RsGral!Etiqueta
        '                 .TextMatrix(.Row, 4) = RsGral!Etiqueta
        '                 .Col = 1
        '                 .CellAlignment = 0
        '                 .TextMatrix(.Row, 1) = RsGral!Formula
        '                 .TextMatrix(.Row, 5) = RsGral!Formula
        '                 .TextMatrix(.Row, 2) = RsGral!Columna
        '                 .TextMatrix(.Row, 6) = RsGral!Columna
        '                 .TextMatrix(.Row, 3) = RsGral!Saltos
        '                 .TextMatrix(.Row, 7) = RsGral!Saltos
        '                 .TextMatrix(.Row, 8) = RsGral!Renglon
        '
        '                 .TextMatrix(.Row, 9) = RsGral!Grupo
        '                 .Row = .Row + 1
        '            End With
        '            .MoveNext
        '            If Not .EOF Then
        '                If RsGral!Renglon = 1 Then
        '                    J = J + 1
        '                End If
        '            End If
        '        Loop
        '    End With
        '    For J = 0 To 2
        '        With FlexDetalle(J)
        '            .Row = 1
        '            .Col = 0
        '        End With
        '    Next
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub

    Sub InsertarLinea()
        'On Error GoTo Errores
        Dim I As Integer
        With FlexDetalle(SSTab1.SelectedIndex)
            .AddItem("", .Row)
        End With
Errores:
        If Err.Number <> 0 Then Errores()
    End Sub

    Sub EliminarLinea()
        Dim TotRen As Integer
        Dim blnTransaccion As Boolean
        Dim strRenglon As String
        Dim strEtiqueta As String
        Dim strFormula As String
        Dim strColumna As String
        Dim strSaltos As String
        Dim strGrupo As String
        Dim strTipo As String
        'On Error GoTo Errores
        If optContado.Checked = True Then
            strTipo = "CO"
        ElseIf optCredito.Checked = True Then
            strTipo = "CR"
        End If
        Cnn.BeginTrans()
        blnTransaccion = True
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        With FlexDetalle(SSTab1.SelectedIndex)
            strRenglon = .get_TextMatrix(.Row, 8)
            strEtiqueta = .get_TextMatrix(.Row, 0)
            strFormula = .get_TextMatrix(.Row, 1)
            strColumna = .get_TextMatrix(.Row, 2)
            strSaltos = .get_TextMatrix(.Row, 3)
            strGrupo = .get_TextMatrix(.Row, 9)
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            ModStoredProcedures.PR_IMEConfigTicketVenta(CStr(intCodSucursal), strRenglon, strEtiqueta, strFormula, strColumna, strSaltos, strGrupo, strTipo, C_ELIMINACION, CStr(0))
            Cmd.Execute()
            TotRen = FlexDetalle(SSTab1.SelectedIndex).Rows
            FlexDetalle(SSTab1.SelectedIndex).RemoveItem((FlexDetalle(SSTab1.SelectedIndex).Row))
            FlexDetalle(SSTab1.SelectedIndex).Rows = TotRen
            FlexDetalle_EnterCell(FlexDetalle.Item((SSTab1.SelectedIndex)), New System.EventArgs())
        End With
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Cnn.CommitTrans()
        blnTransaccion = False
Errores:
        If Err.Number <> 0 Then
            If blnTransaccion = True Then Cnn.RollbackTrans()
            Me.Cursor = System.Windows.Forms.Cursors.Default
            ModEstandar.MostrarError()
        End If
    End Sub

    Function ChecaRenglones(ByRef Renglon As Integer, ByRef Index As Integer) As Boolean
        If Renglon = 1 Then
            ChecaRenglones = True
        Else
            If (FlexDetalle(Index).get_TextMatrix(Renglon - 1, 0) = "" And FlexDetalle(Index).get_TextMatrix(Renglon - 1, 1) = "") Or (FlexDetalle(Index).get_TextMatrix(Renglon - 1, 2) = "" Or FlexDetalle(Index).get_TextMatrix(Renglon - 1, 3) = "") Then
                ChecaRenglones = False
            Else
                ChecaRenglones = True
            End If
        End If
    End Function

    'Sub MSHFlexEdit(ByRef MSHFlexGrid As Control, Edt As Control, KeyAscii As Integer)
    ''On Error GoTo Errores
    '   ' Usar el carácter escrito.
    '   Select Case KeyAscii
    '   ' Un espacio significa modificar el texto actual.
    '   Case 0 To 32
    '      Edt = MSHFlexGrid
    '      'cRenAnt = MSHFlexGrid.Row
    '
    '   ' Otro carácter reemplaza el texto actual.
    '   Case Else
    '        Edt = Chr(KeyAscii)
    '        Edt.SelStart = 1
    '   End Select
    '   ' Mostrar Edt en la posición correcta.
    '    Edt.Move MSHFlexGrid.Left + 90 + MSHFlexGrid.CellLeft, _
    ''    MSHFlexGrid.Top + 320 + MSHFlexGrid.CellTop, _
    ''    MSHFlexGrid.CellWidth + 5
    '
    '    Edt.Visible = True
    '    Edt.Enabled = True
    '    'Y hacer que funcione.
    '    Edt.SetFocus
    'Errores:
    '    If Err.Number <> 0 Then Errores
    ''    Resume
    'End Sub

    Sub MSHFlexEdit(ByRef MSHFlexGrid As System.Windows.Forms.Control, ByRef Edt As System.Windows.Forms.Control, ByRef KeyAscii As Integer)
        'On Error GoTo Errores
        ' Usar el carácter escrito.
        Select Case KeyAscii
            ' Un espacio significa modificar el texto actual.
            Case 0 To 32
                Edt = MSHFlexGrid
                'cRenAnt = MSHFlexGrid.Row

                ' Otro carácter reemplaza el texto actual.
            Case Else
                'Edt = Chr(KeyAscii)
                'Edt.SelStart = 1
        End Select
        ' Mostrar Edt en la posición correcta.
        'If MSHFlexGrid.Col = 0 Then
        ' Edt.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(MSHFlexGrid.Left) + 210 + MSHFlexGrid.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(MSHFlexGrid.Top) + 1310 + MSHFlexGrid.CellTop), VB6.TwipsToPixelsX(MSHFlexGrid.CellWidth + 35), 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
        'ElseIf MSHFlexGrid.Col = 2 Then
        'Edt.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(MSHFlexGrid.Left) + 100 + MSHFlexGrid.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(MSHFlexGrid.Top) + 380 + MSHFlexGrid.CellTop), VB6.TwipsToPixelsX(MSHFlexGrid.CellWidth + 10), 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
        'ElseIf MSHFlexGrid.Col = 3 Then
        'Edt.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(MSHFlexGrid.Left) + 230 + MSHFlexGrid.CellLeft), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(MSHFlexGrid.Top) + 1310 + MSHFlexGrid.CellTop), VB6.TwipsToPixelsX(MSHFlexGrid.CellWidth + 20), 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y Or Windows.Forms.BoundsSpecified.Width)
        'End If
        Edt.Visible = True
        Edt.Enabled = True
        'Y hacer que funcione.
        Edt.Focus()
Errores:
        If Err.Number <> 0 Then Errores()
    End Sub

    Sub InicializaVariables()
        mblnSALIR = False
        mblnSaliryGrabar = False
        mblnPierdeFoco = False
    End Sub

    Sub Encabezado0(ByRef iNumFlex As Integer)
        If (Not isLoad) Then
            Exit Sub
        End If

        With FlexDetalle(iNumFlex)
            .Row = 0
            .Col = 0
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(0, 0, 1890)
            .Text = "Etiqueta"
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(1, 0, 5000)
            .Text = "Fórmula"
            .Col = 2
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(2, 0, 1000)
            .Text = "Columna"
            .Col = 3
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(3, 0, 990)
            .Text = "Saltos"
            .Col = 4
            .set_ColWidth(4, 0, 0)
            .Col = 5
            .set_ColWidth(5, 0, 0)
            .Col = 6
            .set_ColWidth(6, 0, 0)
            .Col = 7
            .set_ColWidth(7, 0, 0)
            .Col = 8
            .set_ColWidth(8, 0, 0)
            .Col = 9
            .set_ColWidth(9, 0, 0)
            .Row = 1
            .Col = 0
        End With
    End Sub

    Sub Encabezado1(ByRef iNumFlex As Integer)
        If (Not isLoad) Then
            Exit Sub
        End If

        With FlexDetalle(iNumFlex)
            .Row = 0
            .Col = 0
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(0, 0, 1890)
            .Text = "Etiqueta"
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(1, 0, 5000)
            .Text = "Fórmula"
            .Col = 2
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(2, 0, 1000)
            .Text = "Columna"
            .Col = 3
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(3, 0, 990)
            .Text = "Saltos"
            .Col = 4
            .set_ColWidth(4, 0, 0)
            .Col = 5
            .set_ColWidth(5, 0, 0)
            .Col = 6
            .set_ColWidth(6, 0, 0)
            .Col = 7
            .set_ColWidth(7, 0, 0)
            .Col = 8
            .set_ColWidth(8, 0, 0)
            .Col = 9
            .set_ColWidth(9, 0, 0)
            .Row = 1
            .Col = 0
        End With
    End Sub

    Sub Encabezado2(ByRef iNumFlex As Integer)
        If (Not isLoad) Then
            Exit Sub
        End If

        With FlexDetalle(iNumFlex)
            .Row = 0
            .Col = 0
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(0, 0, 1890)
            .Text = "Etiqueta"
            .Col = 1
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(1, 0, 5000)
            .Text = "Fórmula"
            .Col = 2
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(2, 0, 1000)
            .Text = "Columna"
            .Col = 3
            .CellAlignment = 5
            .CellFontBold = True
            .set_ColWidth(3, 0, 990)
            .Text = "Saltos"
            .Col = 4
            .set_ColWidth(4, 0, 0)
            .Col = 5
            .set_ColWidth(5, 0, 0)
            .Col = 6
            .set_ColWidth(6, 0, 0)
            .Col = 7
            .set_ColWidth(7, 0, 0)
            .Col = 8
            .set_ColWidth(8, 0, 0)
            .Col = 9
            .set_ColWidth(9, 0, 0)
            .Row = 1
            .Col = 0
        End With
    End Sub
    Private Sub btnDatosDet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDatosDet.Click
        Dim Index As Integer = btnDatosDet.GetIndex(eventSender)
        'Poner enla Columna de Grupo, Que tipo es
        Dim Formula As String
        FlexDetalle(1).set_TextMatrix(FlexDetalle(1).Row, 9, "D")
        If Index < 10 Then
            If FlexDetalle(1).Col = 1 And ChecaRenglones(FlexDetalle(1).Row, 1) Then
                FlexDetalle(1).CellAlignment = 0
                Formula = FlexDetalle(1).Text & "(" & btnDatosDet(Index).Text & ")"
                If Len(Formula) <= 100 Then ' Si el Tamaño del texto de la formula es mayor de  100, este campño no se agregará.
                    FlexDetalle(1).Text = FlexDetalle(1).Text & "(" & btnDatosDet(Index).Text & ")"
                End If
            End If
        ElseIf Index = 10 Or Index = 11 Or Index = 12 Then
            If FlexDetalle(1).Col = 1 And ChecaRenglones(FlexDetalle(1).Row, 1) And FlexDetalle(1).Text <> "" Then
                FlexDetalle(1).CellAlignment = 0
                Formula = FlexDetalle(1).Text & btnDatosDet(Index).Text
                If Len(Formula) <= 100 Then
                    FlexDetalle(1).Text = FlexDetalle(1).Text & btnDatosDet(Index).Text
                End If
            End If
        ElseIf Index = 13 Then
            If FlexDetalle(1).Col = 1 Then
                FlexDetalle(1).Text = ""
            End If
        End If
    End Sub

    Private Sub btnDatosDet_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDatosDet.Enter
        Dim Index As Integer = btnDatosDet.GetIndex(eventSender)
        Pon_Tool()
        Select Case Index
            Case 13
                SSTab1.TabIndex = 34
        End Select
    End Sub

    Private Sub btnDatosEnc_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDatosEnc.Click
        Dim Index As Integer = btnDatosEnc.GetIndex(eventSender)
        'Poner enla Columna de Grupo, Que tipo es
        Dim Formula As String
        FlexDetalle(0).set_TextMatrix(FlexDetalle(0).Row, 9, "T")
        If Index <= 14 Then 'Valores para la Formula
            If FlexDetalle(0).Col = 1 And ChecaRenglones(FlexDetalle(0).Row, 0) Then
                FlexDetalle(0).CellAlignment = 0
                Formula = FlexDetalle(0).Text & "(" & btnDatosEnc(Index).Text & ")"
                If Len(Formula) <= 100 Then ' Si el Tamaño del texto de la formula es mayor de  100, este campño no se agregará.
                    FlexDetalle(0).Text = FlexDetalle(0).Text & "(" & btnDatosEnc(Index).Text & ")"
                End If
            End If
        ElseIf Index = 15 Then  'Signo + para concatenar
            If FlexDetalle(0).Col = 1 And ChecaRenglones(FlexDetalle(0).Row, 0) And FlexDetalle(0).Text <> "" Then
                FlexDetalle(0).CellAlignment = 0
                Formula = FlexDetalle(0).Text & btnDatosEnc(Index).Text
                If Len(Formula) <= 100 Then ' Si el Tamaño del texto de la formula es mayor de  100, este campño no se agregará.
                    FlexDetalle(0).Text = FlexDetalle(0).Text & btnDatosEnc(Index).Text
                End If
            End If
        ElseIf Index = 16 Then  'Limpiar.
            If FlexDetalle(0).Col = 1 Then
                FlexDetalle(0).Text = ""
            End If
        End If
    End Sub

    Private Sub btnDatosEnc_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDatosEnc.Enter
        Dim Index As Integer = btnDatosEnc.GetIndex(eventSender)
        Pon_Tool()
        Select Case Index
            Case 15
                SSTab1.TabIndex = 19
        End Select
    End Sub

    Private Sub btnDatosTot_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDatosTot.Click
        Dim Index As Integer = btnDatosTot.GetIndex(eventSender)
        'Poner enla Columna de Grupo, Que tipo es
        Dim Formula As String
        FlexDetalle(2).set_TextMatrix(FlexDetalle(2).Row, 9, "T")
        If Index <= 12 Or (Index >= 17 And Index <= 30) Or Index = 35 Then 'Valores par ala Fórmula
            If FlexDetalle(2).Col = 1 And ChecaRenglones(FlexDetalle(2).Row, 2) Then
                FlexDetalle(2).CellAlignment = 0
                Formula = FlexDetalle(2).Text & "(" & btnDatosTot(Index).Text & ")"
                If Len(Formula) <= 100 Then
                    FlexDetalle(2).Text = FlexDetalle(2).Text & "(" & btnDatosTot(Index).Text & ")"
                End If
            End If
        ElseIf Index = 13 Or Index = 14 Or Index = 15 Or Index = 31 Or Index = 32 Or Index = 33 Then
            'Operadores para operaciones aritméticas
            If FlexDetalle(2).Col = 1 And ChecaRenglones(FlexDetalle(2).Row, 2) And FlexDetalle(2).Text <> "" Then
                FlexDetalle(2).CellAlignment = 0
                Formula = FlexDetalle(2).Text & btnDatosTot(Index).Text
                If Len(Formula) <= 100 Then
                    FlexDetalle(2).Text = FlexDetalle(2).Text & btnDatosTot(Index).Text
                End If
            End If
        ElseIf Index = 16 Or Index = 34 Then
            'Limpiar
            If FlexDetalle(2).Col = 1 Then
                FlexDetalle(2).Text = ""
            End If
        End If
    End Sub

    Private Sub btnDatosTot_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDatosTot.Enter
        Dim Index As Integer = btnDatosTot.GetIndex(eventSender)
        Pon_Tool()
        Select Case Index
            Case 16
                SSTab1.TabIndex = 52
        End Select
    End Sub

    Private Sub btnPrueba_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPrueba.Click
        ModCorporativo.TicketVentaPrueba(IIf((optContado.Checked = True), "CO", "CR"), intCodSucursal)
    End Sub

    Private Sub btnPrueba_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPrueba.Enter
        Pon_Tool()
    End Sub

    Private Sub dbcSucursales_CursorChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.CursorChanged
        If FueraChange = True Then Exit Sub
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> "dbcSucursales" Then
        '    Exit Sub
        'End If
        Nuevo()
        If Trim(dbcSucursales.Text) = "" Then
            auxCodSucursal = 0
        End If
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%'  And TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCChange(gStrSql, tecla)
        intCodSucursal = 0
        If dbcSucursales.SelectedItem <> "" Then
            Call dbcSucursales_Leave(dbcSucursales, New System.EventArgs())
        End If
        'mblnNuevo = True
    End Sub

    Private Sub dbcSucursales_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Enter
        'If System.Windows.Forms.Form.ActiveForm.ActiveControl.Name <> dbcSucursales.Name Then Exit Sub
        Pon_Tool()
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen Where TipoAlmacen ='P' ORDER BY DescAlmacen"
        DCGotFocus(gStrSql, dbcSucursales)
    End Sub

    Private Sub dbcSucursales_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As KeyEventArgs) Handles dbcSucursales.KeyDown
        tecla = eventArgs.KeyCode
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Escape Then
            mblnSALIR = True
            Me.Close()
        End If
    End Sub

    Private Sub dbcSucursales_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbcSucursales.Leave
        'If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then Exit Sub
        intCodSucursal = 0
        gStrSql = "SELECT CodAlmacen,LTRIM(RTRIM(DescAlmacen)) as DescAlmacen FROM CatAlmacen WHERE DescAlmacen LIKE '" & Trim(dbcSucursales.Text) & "%' and  TipoAlmacen ='P'  ORDER BY DescAlmacen"
        DCLostFocus(dbcSucursales, gStrSql, intCodSucursal)
        If intCodSucursal = 0 Then auxCodSucursal = 0
        If intCodSucursal = auxCodSucursal Then Exit Sub
        If intCodSucursal <> 0 Then
            auxCodSucursal = intCodSucursal
            ValidarSucursalyMostrarDatos()
        End If
    End Sub

    Private Sub FlexDetalle_DblClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FlexDetalle.DblClick
        'Dim Index As Integer = FlexDetalle.GetIndex(eventSender)
        'flexDetalle_KeyPressEvent(FlexDetalle.Item(Index), New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
    End Sub

    Private Sub flexDetalle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FlexDetalle.Enter
        'Dim Index As Integer = FlexDetalle.GetIndex(eventSender)
        Pon_Tool()
        'lblFormula(Index).Text = FlexDetalle(Index).get_TextMatrix(FlexDetalle(Index).Row, 1)
    End Sub

    Private Sub FlexDetalle_EnterCell(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FlexDetalle.EnterCell
        Dim Index As Integer = FlexDetalle.GetIndex(eventSender)
        'On Error GoTo Errores
        With FlexDetalle(Index)
            lblFormula(Index).Text = .get_TextMatrix(.Row, 1)
        End With
Errores:
        If Err.Number <> 0 Then Errores()
    End Sub

    Private Sub flexDetalle_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyDownEvent) Handles FlexDetalle.KeyDownEvent
        Dim Index As Integer = FlexDetalle.GetIndex(eventSender)
        Select Case Index
            Case 1
                If eventArgs.keyCode = System.Windows.Forms.Keys.Escape Then
                    SSTab1.Focus()
                End If
            Case 2
                If eventArgs.keyCode = System.Windows.Forms.Keys.Escape Then
                    SSTab1.Focus()
                End If
        End Select
        If eventArgs.keyCode = System.Windows.Forms.Keys.Delete Then
            EliminarLinea()
        End If
        If eventArgs.keyCode = System.Windows.Forms.Keys.Insert Then
            InsertarLinea()
        End If
    End Sub

    Private Sub flexDetalle_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent) Handles FlexDetalle.KeyPressEvent
        Dim Index As Integer = FlexDetalle.GetIndex(eventSender)
        If SSTab1.SelectedIndex = 0 Then
            Index = 0
        ElseIf SSTab1.SelectedIndex = 1 Then
            Index = 1
        ElseIf SSTab1.SelectedIndex = 2 Then
            Index = 2
        End If
        With FlexDetalle(Index)
            '''Para que en la columna de porcentage
            '''no deje capturar caracteres sino solo numeros
            If .Col = 0 Then
                ModEstandar.gp_CampoMayusculas(eventArgs.keyAscii)
            End If
            '''si ya se capturo algo entonces se edita el grid
            '''ya sea con numeros, letras o enter
            If eventArgs.keyAscii <> 0 And eventArgs.keyAscii <> System.Windows.Forms.Keys.Escape Then
                If (.Row > 1) Then
                    '''de tal modo que si el renglon es mayor que 1
                    '''y si un renglon antes del renglon actual esta vacio,
                    '''el renlgon actual no se editará
                    If (Trim(.get_TextMatrix(.Row - 1, 0)) = "" And Trim(.get_TextMatrix(.Row - 1, 1)) = "") Or (Trim(.get_TextMatrix(.Row - 1, 2)) = "" Or Trim(.get_TextMatrix(.Row - 1, 3)) = "") Then
                        .Focus()
                        Exit Sub
                    End If
                End If
                If .Col = 0 Then
                    MSHFlexEdit(FlexDetalle(Index), TxtEtiqueta, eventArgs.keyAscii)
                ElseIf .Col = 1 Then
                    .Col = .Col + 1
                    MSHFlexEdit(FlexDetalle(Index), TxtColumna, eventArgs.keyAscii)
                ElseIf .Col = 2 Then
                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                    MSHFlexEdit(FlexDetalle(Index), TxtColumna, eventArgs.keyAscii)
                ElseIf .Col = 3 Then
                    If eventArgs.keyAscii < 48 Or eventArgs.keyAscii > 57 Then eventArgs.keyAscii = 0
                    MSHFlexEdit(FlexDetalle(Index), TxtSaltos, eventArgs.keyAscii)
                End If
                If Len(Trim(TxtEtiqueta.Text)) = 1 Then
                    'System.Windows.Forms.SendKeys.Send("{right}")
                End If
            End If
        End With
    End Sub

    Private Sub flexDetalle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FlexDetalle.Leave
        'Dim Index As Integer = FlexDetalle.GetIndex(eventSender)
        'FlexDetalle(Index).FocusRect = MSHierarchicalFlexGridLib.FocusRectSettings.flexFocusLight
        'lblFormula(Index).Text = ""
    End Sub

    Private Sub frmPVConfigTicketVenta_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        Me.BringToFront()
    End Sub

    Private Sub frmPVConfigTicketVenta_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
    End Sub

    Private Sub frmPVConfigTicketVenta_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                ModEstandar.AvanzarTab(Me)
            Case System.Windows.Forms.Keys.Escape
                ModEstandar.RetrocederTab(Me)
        End Select
    End Sub

    Private Sub frmPVConfigTicketVenta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        If KeyAscii = 39 Then KeyAscii = 180 'Transforma apostrofe en acento
        KeyAscii = ModEstandar.gp_CampoMayusculas(KeyAscii) 'Convierte la letra a mayuscula
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmPVConfigTicketVenta_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        isLoad = True
        '    LlenarDatosTicket
        Encabezado0(0)
        Encabezado1(1)
        Encabezado2(2)
        ModEstandar.ActivaMenu(C_ACTIVADO, C_ACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_ACTIVADO)
        ModEstandar.CentrarForma(Me)
        SSTab1.SelectedIndex = 0
    End Sub

    Sub LlenarDatosTicket()
        Dim strTipo As String
        If optContado.Checked = True Then
            strTipo = "CO"
        ElseIf optCredito.Checked = True Then
            strTipo = "CR"
        End If
        gStrSql = "SELECT * FROM ConfigTicketVenta Where Tipo= '" & strTipo & "' and CodAlmacen = " & intCodSucursal & " order by CodAlmacen, Grupo, Renglon "
        ModEstandar.BorraCmd()
        Cmd.CommandText = "dbo.Up_Select_Datos"
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
        Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 800, gStrSql))
        RsGral = Cmd.Execute
        InicializaVariables()
        '''SSTab1.Tab = 0
        Encabezado0(0)
        Encabezado1(1)
        Encabezado2(2)
        Nuevo()
        If RsGral.RecordCount > 0 Then
            LlenaDatos()
            mblnNuevo = False
        Else
            mblnNuevo = True
        End If
        RsGral.Close()
    End Sub

    Private Sub frmPVConfigTicketVenta_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Si se decea cerrar la forma y esta se encuentra minimisada esta se restaurara
        ModEstandar.RestaurarForma(Me, False)
        'Si se cierra el formulario y existio algun cambio en el registro se
        'informa al usuario del cabio y si desea guardar el registro, ya sea
        'que sea nuevo o un registro modificado
        If Not mblnSALIR And Not mblnSaliryGrabar Then
            If Cambios() = True Then
                Select Case MsgBox(C_msgGUARDAR, MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, gstrNombCortoEmpresa)
                    Case MsgBoxResult.Yes 'Guardar el registro
                        If Guardar() = False Then
                            Cancel = 1
                        End If
                    Case MsgBoxResult.No 'No hace nada y permite el cierre del formulario
                    Case MsgBoxResult.Cancel 'Cancela el cierre del formulario sin guardar
                        Cancel = 1
                End Select
            End If
        Else
            If mblnSaliryGrabar Then
                Cancel = 0
                Exit Sub
            End If
            Select Case MsgBox(C_msgSALIR, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2, gstrNombCortoEmpresa)
                Case MsgBoxResult.Yes
                    Cancel = 0
                Case MsgBoxResult.No
                    mblnSALIR = False
                    SSTab1.Focus()
                    Cancel = 1
            End Select
        End If
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmPVConfigTicketVenta_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ModEstandar.ActivaMenu(C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO, C_DESACTIVADO)
        ModEstandar.LimpiaDescBarraEstado()
        'Me = Nothing
    End Sub

    Private Sub Option2_Click()

    End Sub

    Private Sub optContado_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optContado.CheckedChanged
        If eventSender.Checked Then
            If optContado.Checked = True Then
                fraTotalesContado.Visible = True
                fraTotalesCredito.Visible = False
            End If
            LlenarDatosTicket()
        End If
    End Sub

    Private Sub optCredito_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCredito.CheckedChanged
        If eventSender.Checked Then
            If optCredito.Checked = True Then
                fraTotalesCredito.Visible = True
                fraTotalesContado.Visible = False
            End If
            LlenarDatosTicket()
        End If
    End Sub

    Private Sub SSTab1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTab1.Enter
        Pon_Tool()
        If SSTab1.SelectedIndex = 0 Then
            SSTab1.TabIndex = 0
        ElseIf SSTab1.SelectedIndex = 1 Then
            SSTab1.TabIndex = 19
        ElseIf SSTab1.SelectedIndex = 2 Then
            SSTab1.TabIndex = 34
        End If
    End Sub

    Private Sub SSTab1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles SSTab1.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return And SSTab1.SelectedIndex = 0 Then
            TxtCols.Focus()
        ElseIf KeyCode = System.Windows.Forms.Keys.Return And SSTab1.SelectedIndex = 1 Then
            FlexDetalle(1).Focus()
        ElseIf KeyCode = System.Windows.Forms.Keys.Return And SSTab1.SelectedIndex = 2 Then
            FlexDetalle(2).Focus()
            '    ElseIf KeyCode = vbKeyEscape And SSTab1.Tab = 0 Then
            '        mblnSALIR = True
            '        Unload Me
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape And SSTab1.SelectedIndex = 1 Then
            SSTab1.SelectedIndex = 0
            btnDatosEnc(15).Focus()
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape And SSTab1.SelectedIndex = 2 Then
            SSTab1.SelectedIndex = 1
            btnDatosDet(13).Focus()
        End If
    End Sub

    Private Sub SSTab1_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles SSTab1.KeyUp
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        If (KeyCode = System.Windows.Forms.Keys.Left Or KeyCode = System.Windows.Forms.Keys.Right) And SSTab1.SelectedIndex = 0 Then
            SSTab1.TabIndex = 0
        ElseIf (KeyCode = System.Windows.Forms.Keys.Left Or KeyCode = System.Windows.Forms.Keys.Right) And SSTab1.SelectedIndex = 1 Then
            SSTab1.TabIndex = 19
        ElseIf (KeyCode = System.Windows.Forms.Keys.Left Or KeyCode = System.Windows.Forms.Keys.Right) And SSTab1.SelectedIndex = 2 Then
            SSTab1.TabIndex = 34
        End If
    End Sub

    Private Sub TxtCols_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtCols.Enter
        SSTab1.TabIndex = 0
        Pon_Tool()
    End Sub

    Private Sub TxtCols_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtCols.Leave
        If Val(TxtCols.Text) > 255 Then
            MsgBox("La Longitud Maxima es de 255.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
            TxtCols.Text = CStr(0)
        End If
    End Sub

    Private Sub TxtColumna_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtColumna.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        With FlexDetalle(SSTab1.SelectedIndex)
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    TxtColumna.Visible = False
                    TxtColumna.Text = ""
                    .Focus()
                Case System.Windows.Forms.Keys.Return
                    If Val(TxtColumna.Text) <= Val(TxtCols.Text) Then
                        .set_TextMatrix(.Row, .Col, TxtColumna.Text)
                        TxtColumna.Visible = False
                        TxtColumna.Text = ""
                        .Col = .Col + 1
                        mblnPierdeFoco = True
                        flexDetalle_KeyPressEvent(FlexDetalle.Item((SSTab1.SelectedIndex)), New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
                    Else
                        MsgBox("La Columna Especificada No Debe Excecer al Número de Cólumnas.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        TxtColumna.Text = ""
                    End If
            End Select
        End With
    End Sub

    Private Sub TxtColumna_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtColumna.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub TxtColumna_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtColumna.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        If Not mblnPierdeFoco Then
            TxtColumna_KeyDown(TxtColumna, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
        Else
            mblnPierdeFoco = False
        End If
        TxtColumna.Visible = False
    End Sub

    Private Sub TxtEtiqueta_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtEtiqueta.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        With FlexDetalle(SSTab1.SelectedIndex)
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    TxtEtiqueta.Visible = False
                    TxtEtiqueta.Text = ""
                    .Focus()
                Case System.Windows.Forms.Keys.Return
                    .set_TextMatrix(.Row, .Col, TxtEtiqueta.Text)
                    TxtEtiqueta.Visible = False
                    TxtEtiqueta.Text = ""
                    .Col = .Col + 1
                    .Focus()
            End Select
        End With
    End Sub

    Private Sub TxtEtiqueta_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtEtiqueta.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoAlfanumerico(KeyAscii, "!""#$%&/()=?'¿¡,;.:-_{}[]@*\+*<>")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub TxtEtiqueta_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtEtiqueta.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        TxtEtiqueta_KeyDown(TxtEtiqueta, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
        TxtEtiqueta.Visible = False
    End Sub

    Private Sub TxtSaltos_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtSaltos.KeyDown
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000
        With FlexDetalle(SSTab1.SelectedIndex)
            Select Case KeyCode
                Case System.Windows.Forms.Keys.Escape
                    TxtSaltos.Visible = False
                    TxtSaltos.Text = ""
                    .Focus()
                Case System.Windows.Forms.Keys.Return
                    If Val(TxtSaltos.Text) <= 255 Then
                        .set_TextMatrix(.Row, .Col, TxtSaltos.Text)
                        TxtSaltos.Visible = False
                        TxtSaltos.Text = ""
                        If .Row < .Rows - 1 Then
                            .Row = .Row + 1
                            .Col = 0
                            mblnPierdeFoco = True
                            flexDetalle_KeyPressEvent(FlexDetalle.Item((SSTab1.SelectedIndex)), New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
                            Exit Sub
                        ElseIf .Row = .Rows - 1 Then
                            .Rows = .Rows + 1
                            .Row = .Row + 1
                            .Col = 0
                            mblnPierdeFoco = True
                            flexDetalle_KeyPressEvent(FlexDetalle.Item((SSTab1.SelectedIndex)), New AxMSHierarchicalFlexGridLib.DMSHFlexGridEvents_KeyPressEvent(System.Windows.Forms.Keys.Return))
                        End If
                    Else
                        MsgBox("La Longitud Maxima es de 255.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, gstrNombCortoEmpresa)
                        TxtSaltos.Text = ""
                    End If
            End Select
        End With
    End Sub

    Private Sub TxtSaltos_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtSaltos.KeyPress
        Dim KeyAscii As Integer = Asc(eventArgs.KeyChar)
        ModEstandar.gp_CampoNumerico(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub TxtSaltos_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtSaltos.Leave
        If System.Windows.Forms.Form.ActiveForm.Name <> Me.Name Then
            Exit Sub
        End If
        If Not mblnPierdeFoco Then
            TxtSaltos_KeyDown(TxtSaltos, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape Or 0 * &H10000))
        Else
            mblnPierdeFoco = False
        End If
        TxtSaltos.Visible = False
    End Sub

    'Sub Limpiar()
    '    FlexDetalle(0).Clear
    '    FlexDetalle(1).Clear
    '    FlexDetalle(2).Clear
    '
    '    Encabezado (0)
    '    Encabezado (1)
    '    Encabezado (2)
    'End Sub

    Sub ValidarSucursalyMostrarDatos()
        'On Error GoTo Merr
        If intCodSucursal = 0 Then
            MsgBox("El Código de la Sucursal no existe." & vbNewLine & "Verifique Por Favor.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, gstrNombCortoEmpresa)
            dbcSucursales.Focus()
            Exit Sub
        Else
            LlenarDatosTicket()
        End If
Merr:
        If Err.Number <> 0 Then ModEstandar.MostrarError()
    End Sub


    Public Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPVConfigTicketVenta))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TxtSaltos = New System.Windows.Forms.TextBox()
        Me.TxtEtiqueta = New System.Windows.Forms.TextBox()
        Me.TxtCols = New System.Windows.Forms.TextBox()
        Me.TxtColumna = New System.Windows.Forms.TextBox()
        Me._Label1_0 = New System.Windows.Forms.Label()
        Me.chkAplicarSucursales = New System.Windows.Forms.CheckBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.optCredito = New System.Windows.Forms.RadioButton()
        Me.optContado = New System.Windows.Forms.RadioButton()
        Me.SSTab1 = New System.Windows.Forms.TabControl()
        Me._SSTab1_TabPage0 = New System.Windows.Forms.TabPage()
        Me._Marco_1 = New System.Windows.Forms.GroupBox()
        Me.btnPrueba = New System.Windows.Forms.Button()
        Me._FlexDetalle_0 = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._lblFormula_0 = New System.Windows.Forms.Label()
        Me._Label1_1 = New System.Windows.Forms.Label()
        Me._Marco_0 = New System.Windows.Forms.GroupBox()
        Me._btnDatosEnc_14 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_16 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_15 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_13 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_12 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_11 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_10 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_9 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_8 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_7 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_6 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_5 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_4 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_3 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_1 = New System.Windows.Forms.Button()
        Me._btnDatosEnc_0 = New System.Windows.Forms.Button()
        Me._SSTab1_TabPage1 = New System.Windows.Forms.TabPage()
        Me._Marco_3 = New System.Windows.Forms.GroupBox()
        Me._btnDatosDet_13 = New System.Windows.Forms.Button()
        Me._btnDatosDet_12 = New System.Windows.Forms.Button()
        Me._btnDatosDet_11 = New System.Windows.Forms.Button()
        Me._btnDatosDet_10 = New System.Windows.Forms.Button()
        Me._btnDatosDet_7 = New System.Windows.Forms.Button()
        Me._btnDatosDet_6 = New System.Windows.Forms.Button()
        Me._btnDatosDet_5 = New System.Windows.Forms.Button()
        Me._btnDatosDet_4 = New System.Windows.Forms.Button()
        Me._btnDatosDet_3 = New System.Windows.Forms.Button()
        Me._btnDatosDet_2 = New System.Windows.Forms.Button()
        Me._btnDatosDet_1 = New System.Windows.Forms.Button()
        Me._btnDatosDet_0 = New System.Windows.Forms.Button()
        Me._Marco_2 = New System.Windows.Forms.GroupBox()
        Me._FlexDetalle_1 = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._lblFormula_1 = New System.Windows.Forms.Label()
        Me._SSTab1_TabPage2 = New System.Windows.Forms.TabPage()
        Me.fraTotalesContado = New System.Windows.Forms.GroupBox()
        Me._btnDatosTot_12 = New System.Windows.Forms.Button()
        Me._btnDatosTot_16 = New System.Windows.Forms.Button()
        Me._btnDatosTot_14 = New System.Windows.Forms.Button()
        Me._btnDatosTot_15 = New System.Windows.Forms.Button()
        Me._btnDatosTot_13 = New System.Windows.Forms.Button()
        Me._btnDatosTot_11 = New System.Windows.Forms.Button()
        Me._btnDatosTot_10 = New System.Windows.Forms.Button()
        Me._btnDatosTot_9 = New System.Windows.Forms.Button()
        Me._btnDatosTot_8 = New System.Windows.Forms.Button()
        Me._btnDatosTot_7 = New System.Windows.Forms.Button()
        Me._btnDatosTot_6 = New System.Windows.Forms.Button()
        Me._btnDatosTot_5 = New System.Windows.Forms.Button()
        Me._btnDatosTot_4 = New System.Windows.Forms.Button()
        Me._btnDatosTot_3 = New System.Windows.Forms.Button()
        Me._btnDatosTot_2 = New System.Windows.Forms.Button()
        Me._btnDatosTot_35 = New System.Windows.Forms.Button()
        Me._btnDatosTot_1 = New System.Windows.Forms.Button()
        Me._btnDatosTot_0 = New System.Windows.Forms.Button()
        Me.fraTotalesCredito = New System.Windows.Forms.GroupBox()
        Me._btnDatosTot_30 = New System.Windows.Forms.Button()
        Me._btnDatosTot_29 = New System.Windows.Forms.Button()
        Me._btnDatosTot_34 = New System.Windows.Forms.Button()
        Me._btnDatosTot_32 = New System.Windows.Forms.Button()
        Me._btnDatosTot_33 = New System.Windows.Forms.Button()
        Me._btnDatosTot_31 = New System.Windows.Forms.Button()
        Me._btnDatosTot_28 = New System.Windows.Forms.Button()
        Me._btnDatosTot_27 = New System.Windows.Forms.Button()
        Me._btnDatosTot_26 = New System.Windows.Forms.Button()
        Me._btnDatosTot_25 = New System.Windows.Forms.Button()
        Me._btnDatosTot_24 = New System.Windows.Forms.Button()
        Me._btnDatosTot_23 = New System.Windows.Forms.Button()
        Me._btnDatosTot_22 = New System.Windows.Forms.Button()
        Me._btnDatosTot_21 = New System.Windows.Forms.Button()
        Me._btnDatosTot_20 = New System.Windows.Forms.Button()
        Me._btnDatosTot_19 = New System.Windows.Forms.Button()
        Me._btnDatosTot_18 = New System.Windows.Forms.Button()
        Me._btnDatosTot_17 = New System.Windows.Forms.Button()
        Me._Marco_4 = New System.Windows.Forms.GroupBox()
        Me._FlexDetalle_2 = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me._lblFormula_2 = New System.Windows.Forms.Label()
        Me.dbcSucursales = New System.Windows.Forms.ComboBox()
        Me._Label2_1 = New System.Windows.Forms.Label()
        Me._Label2_0 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.FlexDetalle = New AxMSHFlexGridArray.AxMSHFlexGridArray(Me.components)
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Marco = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.btnDatosDet = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.btnDatosEnc = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.btnDatosTot = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.lblFormula = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Frame1.SuspendLayout()
        Me.SSTab1.SuspendLayout()
        Me._SSTab1_TabPage0.SuspendLayout()
        Me._Marco_1.SuspendLayout()
        CType(Me._FlexDetalle_0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._Marco_0.SuspendLayout()
        Me._SSTab1_TabPage1.SuspendLayout()
        Me._Marco_3.SuspendLayout()
        Me._Marco_2.SuspendLayout()
        CType(Me._FlexDetalle_1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me._SSTab1_TabPage2.SuspendLayout()
        Me.fraTotalesContado.SuspendLayout()
        Me.fraTotalesCredito.SuspendLayout()
        Me._Marco_4.SuspendLayout()
        CType(Me._FlexDetalle_2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FlexDetalle, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Marco, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDatosDet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDatosEnc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDatosTot, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblFormula, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TxtSaltos
        '
        Me.TxtSaltos.AcceptsReturn = True
        Me.TxtSaltos.BackColor = System.Drawing.Color.White
        Me.TxtSaltos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtSaltos.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TxtSaltos.Location = New System.Drawing.Point(92, 180)
        Me.TxtSaltos.MaxLength = 3
        Me.TxtSaltos.Name = "TxtSaltos"
        Me.TxtSaltos.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtSaltos.Size = New System.Drawing.Size(63, 20)
        Me.TxtSaltos.TabIndex = 49
        Me.TxtSaltos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.TxtSaltos, "Número de Saltos")
        Me.TxtSaltos.Visible = False
        '
        'TxtEtiqueta
        '
        Me.TxtEtiqueta.AcceptsReturn = True
        Me.TxtEtiqueta.BackColor = System.Drawing.Color.White
        Me.TxtEtiqueta.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtEtiqueta.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TxtEtiqueta.Location = New System.Drawing.Point(28, 176)
        Me.TxtEtiqueta.MaxLength = 25
        Me.TxtEtiqueta.Name = "TxtEtiqueta"
        Me.TxtEtiqueta.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtEtiqueta.Size = New System.Drawing.Size(64, 20)
        Me.TxtEtiqueta.TabIndex = 47
        Me.ToolTip1.SetToolTip(Me.TxtEtiqueta, "Etiqueta a Mostrar")
        Me.TxtEtiqueta.Visible = False
        '
        'TxtCols
        '
        Me.TxtCols.AcceptsReturn = True
        Me.TxtCols.BackColor = System.Drawing.SystemColors.Window
        Me.TxtCols.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtCols.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TxtCols.Location = New System.Drawing.Point(97, 16)
        Me.TxtCols.MaxLength = 3
        Me.TxtCols.Name = "TxtCols"
        Me.TxtCols.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtCols.Size = New System.Drawing.Size(44, 21)
        Me.TxtCols.TabIndex = 10
        Me.TxtCols.Text = "0"
        Me.TxtCols.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.TxtCols, "Numero de Columnas que Contendra el Ticket.")
        '
        'TxtColumna
        '
        Me.TxtColumna.AcceptsReturn = True
        Me.TxtColumna.BackColor = System.Drawing.Color.White
        Me.TxtColumna.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtColumna.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TxtColumna.Location = New System.Drawing.Point(148, 118)
        Me.TxtColumna.MaxLength = 3
        Me.TxtColumna.Name = "TxtColumna"
        Me.TxtColumna.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtColumna.Size = New System.Drawing.Size(63, 20)
        Me.TxtColumna.TabIndex = 48
        Me.TxtColumna.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.TxtColumna, "Columna Donde Inicia el Dato")
        Me.TxtColumna.Visible = False
        '
        '_Label1_0
        '
        Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.ForeColor = System.Drawing.Color.Black
        Me._Label1_0.Location = New System.Drawing.Point(16, 24)
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.Size = New System.Drawing.Size(60, 17)
        Me._Label1_0.TabIndex = 0
        Me._Label1_0.Text = "Sucursal :"
        Me.ToolTip1.SetToolTip(Me._Label1_0, "Nombre de la Farmacia Actual")
        '
        'chkAplicarSucursales
        '
        Me.chkAplicarSucursales.BackColor = System.Drawing.SystemColors.Control
        Me.chkAplicarSucursales.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkAplicarSucursales.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chkAplicarSucursales.Location = New System.Drawing.Point(608, 32)
        Me.chkAplicarSucursales.Name = "chkAplicarSucursales"
        Me.chkAplicarSucursales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkAplicarSucursales.Size = New System.Drawing.Size(176, 23)
        Me.chkAplicarSucursales.TabIndex = 92
        Me.chkAplicarSucursales.Text = "Aplicar esta configuración a todas las sucursales"
        Me.chkAplicarSucursales.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.optCredito)
        Me.Frame1.Controls.Add(Me.optContado)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(304, 8)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(249, 49)
        Me.Frame1.TabIndex = 2
        Me.Frame1.TabStop = False
        Me.Frame1.Text = " Tipo de Ticket "
        '
        'optCredito
        '
        Me.optCredito.BackColor = System.Drawing.SystemColors.Control
        Me.optCredito.Cursor = System.Windows.Forms.Cursors.Default
        Me.optCredito.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optCredito.Location = New System.Drawing.Point(136, 20)
        Me.optCredito.Name = "optCredito"
        Me.optCredito.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optCredito.Size = New System.Drawing.Size(65, 21)
        Me.optCredito.TabIndex = 4
        Me.optCredito.TabStop = True
        Me.optCredito.Text = "Crédito"
        Me.optCredito.UseVisualStyleBackColor = False
        '
        'optContado
        '
        Me.optContado.BackColor = System.Drawing.SystemColors.Control
        Me.optContado.Checked = True
        Me.optContado.Cursor = System.Windows.Forms.Cursors.Default
        Me.optContado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optContado.Location = New System.Drawing.Point(40, 20)
        Me.optContado.Name = "optContado"
        Me.optContado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optContado.Size = New System.Drawing.Size(81, 21)
        Me.optContado.TabIndex = 3
        Me.optContado.TabStop = True
        Me.optContado.Text = "Contado"
        Me.optContado.UseVisualStyleBackColor = False
        '
        'SSTab1
        '
        Me.SSTab1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage0)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage1)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage2)
        Me.SSTab1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTab1.Location = New System.Drawing.Point(8, 62)
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.SelectedIndex = 0
        Me.SSTab1.Size = New System.Drawing.Size(794, 474)
        Me.SSTab1.TabIndex = 5
        '
        '_SSTab1_TabPage0
        '
        Me._SSTab1_TabPage0.Controls.Add(Me._Marco_1)
        Me._SSTab1_TabPage0.Controls.Add(Me._Marco_0)
        Me._SSTab1_TabPage0.Controls.Add(Me.TxtColumna)
        Me._SSTab1_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage0.Name = "_SSTab1_TabPage0"
        Me._SSTab1_TabPage0.Size = New System.Drawing.Size(786, 448)
        Me._SSTab1_TabPage0.TabIndex = 0
        Me._SSTab1_TabPage0.Text = "Encabezado"
        '
        '_Marco_1
        '
        Me._Marco_1.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_1.Controls.Add(Me.btnPrueba)
        Me._Marco_1.Controls.Add(Me._FlexDetalle_0)
        Me._Marco_1.Controls.Add(Me.TxtCols)
        Me._Marco_1.Controls.Add(Me._lblFormula_0)
        Me._Marco_1.Controls.Add(Me._Label1_1)
        Me._Marco_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_1.Location = New System.Drawing.Point(9, 27)
        Me._Marco_1.Name = "_Marco_1"
        Me._Marco_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_1.Size = New System.Drawing.Size(630, 429)
        Me._Marco_1.TabIndex = 40
        Me._Marco_1.TabStop = False
        '
        'btnPrueba
        '
        Me.btnPrueba.BackColor = System.Drawing.SystemColors.Control
        Me.btnPrueba.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnPrueba.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPrueba.Location = New System.Drawing.Point(527, 16)
        Me.btnPrueba.Name = "btnPrueba"
        Me.btnPrueba.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnPrueba.Size = New System.Drawing.Size(92, 24)
        Me.btnPrueba.TabIndex = 12
        Me.btnPrueba.Text = "&Imprimir Prueba"
        Me.btnPrueba.UseVisualStyleBackColor = False
        '
        '_FlexDetalle_0
        '
        Me._FlexDetalle_0.DataSource = Nothing
        Me._FlexDetalle_0.Location = New System.Drawing.Point(9, 72)
        Me._FlexDetalle_0.Name = "_FlexDetalle_0"
        Me._FlexDetalle_0.OcxState = CType(resources.GetObject("_FlexDetalle_0.OcxState"), System.Windows.Forms.AxHost.State)
        Me._FlexDetalle_0.Size = New System.Drawing.Size(612, 346)
        Me._FlexDetalle_0.TabIndex = 11
        '
        '_lblFormula_0
        '
        Me._lblFormula_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer))
        Me._lblFormula_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblFormula_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormula_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblFormula_0.Location = New System.Drawing.Point(9, 43)
        Me._lblFormula_0.Name = "_lblFormula_0"
        Me._lblFormula_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormula_0.Size = New System.Drawing.Size(518, 22)
        Me._lblFormula_0.TabIndex = 42
        Me._lblFormula_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_Label1_1
        '
        Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_1.Location = New System.Drawing.Point(11, 20)
        Me._Label1_1.Name = "_Label1_1"
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.Size = New System.Drawing.Size(84, 16)
        Me._Label1_1.TabIndex = 41
        Me._Label1_1.Text = "Núm. Columnas:"
        '
        '_Marco_0
        '
        Me._Marco_0.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_14)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_16)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_15)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_13)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_12)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_11)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_10)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_9)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_8)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_7)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_6)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_5)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_4)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_3)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_1)
        Me._Marco_0.Controls.Add(Me._btnDatosEnc_0)
        Me._Marco_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_0.Location = New System.Drawing.Point(648, 27)
        Me._Marco_0.Name = "_Marco_0"
        Me._Marco_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_0.Size = New System.Drawing.Size(135, 429)
        Me._Marco_0.TabIndex = 43
        Me._Marco_0.TabStop = False
        '
        '_btnDatosEnc_14
        '
        Me._btnDatosEnc_14.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_14.Location = New System.Drawing.Point(1, 267)
        Me._btnDatosEnc_14.Name = "_btnDatosEnc_14"
        Me._btnDatosEnc_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_14.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_14.TabIndex = 89
        Me._btnDatosEnc_14.Text = "Tipo de Cambio"
        Me._btnDatosEnc_14.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_16
        '
        Me._btnDatosEnc_16.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_16.Location = New System.Drawing.Point(67, 407)
        Me._btnDatosEnc_16.Name = "_btnDatosEnc_16"
        Me._btnDatosEnc_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_16.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosEnc_16.TabIndex = 27
        Me._btnDatosEnc_16.Text = "Limpiar"
        Me._btnDatosEnc_16.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_15
        '
        Me._btnDatosEnc_15.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_15.Location = New System.Drawing.Point(1, 407)
        Me._btnDatosEnc_15.Name = "_btnDatosEnc_15"
        Me._btnDatosEnc_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_15.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosEnc_15.TabIndex = 26
        Me._btnDatosEnc_15.Text = "+"
        Me._btnDatosEnc_15.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_13
        '
        Me._btnDatosEnc_13.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_13.Location = New System.Drawing.Point(1, 247)
        Me._btnDatosEnc_13.Name = "_btnDatosEnc_13"
        Me._btnDatosEnc_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_13.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_13.TabIndex = 25
        Me._btnDatosEnc_13.Text = "Cliente"
        Me._btnDatosEnc_13.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_12
        '
        Me._btnDatosEnc_12.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_12.Location = New System.Drawing.Point(1, 227)
        Me._btnDatosEnc_12.Name = "_btnDatosEnc_12"
        Me._btnDatosEnc_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_12.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_12.TabIndex = 24
        Me._btnDatosEnc_12.Text = "Vendedor"
        Me._btnDatosEnc_12.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_11
        '
        Me._btnDatosEnc_11.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_11.Location = New System.Drawing.Point(1, 207)
        Me._btnDatosEnc_11.Name = "_btnDatosEnc_11"
        Me._btnDatosEnc_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_11.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_11.TabIndex = 23
        Me._btnDatosEnc_11.Text = "Cajero"
        Me._btnDatosEnc_11.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_10
        '
        Me._btnDatosEnc_10.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_10.Location = New System.Drawing.Point(1, 187)
        Me._btnDatosEnc_10.Name = "_btnDatosEnc_10"
        Me._btnDatosEnc_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_10.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_10.TabIndex = 22
        Me._btnDatosEnc_10.Text = "Tipo de Venta"
        Me._btnDatosEnc_10.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_9
        '
        Me._btnDatosEnc_9.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_9.Location = New System.Drawing.Point(1, 167)
        Me._btnDatosEnc_9.Name = "_btnDatosEnc_9"
        Me._btnDatosEnc_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_9.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_9.TabIndex = 21
        Me._btnDatosEnc_9.Text = "Hora"
        Me._btnDatosEnc_9.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_8
        '
        Me._btnDatosEnc_8.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_8.Location = New System.Drawing.Point(1, 147)
        Me._btnDatosEnc_8.Name = "_btnDatosEnc_8"
        Me._btnDatosEnc_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_8.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_8.TabIndex = 20
        Me._btnDatosEnc_8.Text = "Fecha"
        Me._btnDatosEnc_8.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_7
        '
        Me._btnDatosEnc_7.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_7.Location = New System.Drawing.Point(1, 127)
        Me._btnDatosEnc_7.Name = "_btnDatosEnc_7"
        Me._btnDatosEnc_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_7.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_7.TabIndex = 19
        Me._btnDatosEnc_7.Text = "Folio"
        Me._btnDatosEnc_7.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_6
        '
        Me._btnDatosEnc_6.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_6.Location = New System.Drawing.Point(1, 107)
        Me._btnDatosEnc_6.Name = "_btnDatosEnc_6"
        Me._btnDatosEnc_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_6.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_6.TabIndex = 18
        Me._btnDatosEnc_6.Text = "RFC"
        Me._btnDatosEnc_6.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_5
        '
        Me._btnDatosEnc_5.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_5.Location = New System.Drawing.Point(1, 87)
        Me._btnDatosEnc_5.Name = "_btnDatosEnc_5"
        Me._btnDatosEnc_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_5.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_5.TabIndex = 17
        Me._btnDatosEnc_5.Text = "Ciudad Sucursal"
        Me._btnDatosEnc_5.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_4
        '
        Me._btnDatosEnc_4.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_4.Location = New System.Drawing.Point(1, 67)
        Me._btnDatosEnc_4.Name = "_btnDatosEnc_4"
        Me._btnDatosEnc_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_4.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_4.TabIndex = 16
        Me._btnDatosEnc_4.Text = "Dirección Sucursal"
        Me._btnDatosEnc_4.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_3
        '
        Me._btnDatosEnc_3.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_3.Location = New System.Drawing.Point(1, 48)
        Me._btnDatosEnc_3.Name = "_btnDatosEnc_3"
        Me._btnDatosEnc_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_3.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_3.TabIndex = 15
        Me._btnDatosEnc_3.Text = "Sucursal"
        Me._btnDatosEnc_3.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_1
        '
        Me._btnDatosEnc_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_1.Location = New System.Drawing.Point(1, 27)
        Me._btnDatosEnc_1.Name = "_btnDatosEnc_1"
        Me._btnDatosEnc_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_1.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_1.TabIndex = 14
        Me._btnDatosEnc_1.Text = "Dirección Empresa"
        Me._btnDatosEnc_1.UseVisualStyleBackColor = False
        '
        '_btnDatosEnc_0
        '
        Me._btnDatosEnc_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosEnc_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosEnc_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosEnc_0.Location = New System.Drawing.Point(1, 7)
        Me._btnDatosEnc_0.Name = "_btnDatosEnc_0"
        Me._btnDatosEnc_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosEnc_0.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosEnc_0.TabIndex = 13
        Me._btnDatosEnc_0.Text = "Empresa"
        Me._btnDatosEnc_0.UseVisualStyleBackColor = False
        '
        '_SSTab1_TabPage1
        '
        Me._SSTab1_TabPage1.Controls.Add(Me._Marco_3)
        Me._SSTab1_TabPage1.Controls.Add(Me._Marco_2)
        Me._SSTab1_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage1.Name = "_SSTab1_TabPage1"
        Me._SSTab1_TabPage1.Size = New System.Drawing.Size(786, 448)
        Me._SSTab1_TabPage1.TabIndex = 1
        Me._SSTab1_TabPage1.Text = "Detalle"
        '
        '_Marco_3
        '
        Me._Marco_3.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_3.Controls.Add(Me._btnDatosDet_13)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_12)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_11)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_10)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_7)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_6)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_5)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_4)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_3)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_2)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_1)
        Me._Marco_3.Controls.Add(Me._btnDatosDet_0)
        Me._Marco_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_3.Location = New System.Drawing.Point(648, 27)
        Me._Marco_3.Name = "_Marco_3"
        Me._Marco_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_3.Size = New System.Drawing.Size(135, 429)
        Me._Marco_3.TabIndex = 45
        Me._Marco_3.TabStop = False
        '
        '_btnDatosDet_13
        '
        Me._btnDatosDet_13.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_13.Location = New System.Drawing.Point(67, 407)
        Me._btnDatosDet_13.Name = "_btnDatosDet_13"
        Me._btnDatosDet_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_13.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosDet_13.TabIndex = 39
        Me._btnDatosDet_13.Text = "Limpiar"
        Me._btnDatosDet_13.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_12
        '
        Me._btnDatosDet_12.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_12.Location = New System.Drawing.Point(67, 387)
        Me._btnDatosDet_12.Name = "_btnDatosDet_12"
        Me._btnDatosDet_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_12.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosDet_12.TabIndex = 38
        Me._btnDatosDet_12.Text = "x"
        Me._btnDatosDet_12.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_11
        '
        Me._btnDatosDet_11.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_11.Location = New System.Drawing.Point(1, 407)
        Me._btnDatosDet_11.Name = "_btnDatosDet_11"
        Me._btnDatosDet_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_11.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosDet_11.TabIndex = 37
        Me._btnDatosDet_11.Text = "-"
        Me._btnDatosDet_11.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_10
        '
        Me._btnDatosDet_10.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_10.Location = New System.Drawing.Point(1, 387)
        Me._btnDatosDet_10.Name = "_btnDatosDet_10"
        Me._btnDatosDet_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_10.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosDet_10.TabIndex = 36
        Me._btnDatosDet_10.Text = "+"
        Me._btnDatosDet_10.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_7
        '
        Me._btnDatosDet_7.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_7.Location = New System.Drawing.Point(1, 147)
        Me._btnDatosDet_7.Name = "_btnDatosDet_7"
        Me._btnDatosDet_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_7.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosDet_7.TabIndex = 35
        Me._btnDatosDet_7.Text = "% IVA"
        Me._btnDatosDet_7.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_6
        '
        Me._btnDatosDet_6.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_6.Location = New System.Drawing.Point(1, 127)
        Me._btnDatosDet_6.Name = "_btnDatosDet_6"
        Me._btnDatosDet_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_6.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosDet_6.TabIndex = 34
        Me._btnDatosDet_6.Text = "$ IVA"
        Me._btnDatosDet_6.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_5
        '
        Me._btnDatosDet_5.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_5.Location = New System.Drawing.Point(1, 107)
        Me._btnDatosDet_5.Name = "_btnDatosDet_5"
        Me._btnDatosDet_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_5.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosDet_5.TabIndex = 33
        Me._btnDatosDet_5.Text = "$ Descuento"
        Me._btnDatosDet_5.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_4
        '
        Me._btnDatosDet_4.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_4.Location = New System.Drawing.Point(1, 87)
        Me._btnDatosDet_4.Name = "_btnDatosDet_4"
        Me._btnDatosDet_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_4.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosDet_4.TabIndex = 32
        Me._btnDatosDet_4.Text = "% Descuento"
        Me._btnDatosDet_4.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_3
        '
        Me._btnDatosDet_3.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_3.Location = New System.Drawing.Point(1, 67)
        Me._btnDatosDet_3.Name = "_btnDatosDet_3"
        Me._btnDatosDet_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_3.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosDet_3.TabIndex = 31
        Me._btnDatosDet_3.Text = "Precio Público"
        Me._btnDatosDet_3.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_2
        '
        Me._btnDatosDet_2.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_2.Location = New System.Drawing.Point(1, 47)
        Me._btnDatosDet_2.Name = "_btnDatosDet_2"
        Me._btnDatosDet_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_2.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosDet_2.TabIndex = 30
        Me._btnDatosDet_2.Text = "Cantidad"
        Me._btnDatosDet_2.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_1
        '
        Me._btnDatosDet_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_1.Location = New System.Drawing.Point(1, 27)
        Me._btnDatosDet_1.Name = "_btnDatosDet_1"
        Me._btnDatosDet_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_1.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosDet_1.TabIndex = 29
        Me._btnDatosDet_1.Text = "Descripción"
        Me._btnDatosDet_1.UseVisualStyleBackColor = False
        '
        '_btnDatosDet_0
        '
        Me._btnDatosDet_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosDet_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosDet_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosDet_0.Location = New System.Drawing.Point(1, 7)
        Me._btnDatosDet_0.Name = "_btnDatosDet_0"
        Me._btnDatosDet_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosDet_0.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosDet_0.TabIndex = 28
        Me._btnDatosDet_0.Text = "Código"
        Me._btnDatosDet_0.UseVisualStyleBackColor = False
        '
        '_Marco_2
        '
        Me._Marco_2.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_2.Controls.Add(Me._FlexDetalle_1)
        Me._Marco_2.Controls.Add(Me._lblFormula_1)
        Me._Marco_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_2.Location = New System.Drawing.Point(9, 27)
        Me._Marco_2.Name = "_Marco_2"
        Me._Marco_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_2.Size = New System.Drawing.Size(630, 429)
        Me._Marco_2.TabIndex = 44
        Me._Marco_2.TabStop = False
        '
        '_FlexDetalle_1
        '
        Me._FlexDetalle_1.DataSource = Nothing
        Me._FlexDetalle_1.Location = New System.Drawing.Point(9, 71)
        Me._FlexDetalle_1.Name = "_FlexDetalle_1"
        Me._FlexDetalle_1.OcxState = CType(resources.GetObject("_FlexDetalle_1.OcxState"), System.Windows.Forms.AxHost.State)
        Me._FlexDetalle_1.Size = New System.Drawing.Size(612, 346)
        Me._FlexDetalle_1.TabIndex = 7
        '
        '_lblFormula_1
        '
        Me._lblFormula_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer))
        Me._lblFormula_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblFormula_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormula_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblFormula_1.Location = New System.Drawing.Point(9, 43)
        Me._lblFormula_1.Name = "_lblFormula_1"
        Me._lblFormula_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormula_1.Size = New System.Drawing.Size(613, 22)
        Me._lblFormula_1.TabIndex = 6
        Me._lblFormula_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        '_SSTab1_TabPage2
        '
        Me._SSTab1_TabPage2.Controls.Add(Me.fraTotalesContado)
        Me._SSTab1_TabPage2.Controls.Add(Me.fraTotalesCredito)
        Me._SSTab1_TabPage2.Controls.Add(Me._Marco_4)
        Me._SSTab1_TabPage2.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage2.Name = "_SSTab1_TabPage2"
        Me._SSTab1_TabPage2.Size = New System.Drawing.Size(786, 448)
        Me._SSTab1_TabPage2.TabIndex = 2
        Me._SSTab1_TabPage2.Text = "Totales"
        '
        'fraTotalesContado
        '
        Me.fraTotalesContado.BackColor = System.Drawing.SystemColors.Control
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_12)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_16)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_14)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_15)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_13)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_11)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_10)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_9)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_8)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_7)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_6)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_5)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_4)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_3)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_2)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_35)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_1)
        Me.fraTotalesContado.Controls.Add(Me._btnDatosTot_0)
        Me.fraTotalesContado.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTotalesContado.Location = New System.Drawing.Point(648, 27)
        Me.fraTotalesContado.Name = "fraTotalesContado"
        Me.fraTotalesContado.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTotalesContado.Size = New System.Drawing.Size(135, 429)
        Me.fraTotalesContado.TabIndex = 69
        Me.fraTotalesContado.TabStop = False
        '
        '_btnDatosTot_12
        '
        Me._btnDatosTot_12.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_12.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_12.Location = New System.Drawing.Point(1, 267)
        Me._btnDatosTot_12.Name = "_btnDatosTot_12"
        Me._btnDatosTot_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_12.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_12.TabIndex = 90
        Me._btnDatosTot_12.Text = "Moneda"
        Me._btnDatosTot_12.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_16
        '
        Me._btnDatosTot_16.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_16.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_16.Location = New System.Drawing.Point(67, 407)
        Me._btnDatosTot_16.Name = "_btnDatosTot_16"
        Me._btnDatosTot_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_16.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosTot_16.TabIndex = 80
        Me._btnDatosTot_16.Text = "Limpiar"
        Me._btnDatosTot_16.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_14
        '
        Me._btnDatosTot_14.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_14.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_14.Location = New System.Drawing.Point(67, 387)
        Me._btnDatosTot_14.Name = "_btnDatosTot_14"
        Me._btnDatosTot_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_14.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosTot_14.TabIndex = 85
        Me._btnDatosTot_14.Text = "x"
        Me._btnDatosTot_14.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_15
        '
        Me._btnDatosTot_15.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_15.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_15.Location = New System.Drawing.Point(1, 407)
        Me._btnDatosTot_15.Name = "_btnDatosTot_15"
        Me._btnDatosTot_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_15.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosTot_15.TabIndex = 84
        Me._btnDatosTot_15.Text = "-"
        Me._btnDatosTot_15.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_13
        '
        Me._btnDatosTot_13.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_13.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_13.Location = New System.Drawing.Point(1, 387)
        Me._btnDatosTot_13.Name = "_btnDatosTot_13"
        Me._btnDatosTot_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_13.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosTot_13.TabIndex = 83
        Me._btnDatosTot_13.Text = "+"
        Me._btnDatosTot_13.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_11
        '
        Me._btnDatosTot_11.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_11.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_11.Location = New System.Drawing.Point(1, 247)
        Me._btnDatosTot_11.Name = "_btnDatosTot_11"
        Me._btnDatosTot_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_11.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_11.TabIndex = 82
        Me._btnDatosTot_11.Text = "Total Piezas"
        Me._btnDatosTot_11.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_10
        '
        Me._btnDatosTot_10.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_10.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_10.Location = New System.Drawing.Point(1, 227)
        Me._btnDatosTot_10.Name = "_btnDatosTot_10"
        Me._btnDatosTot_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_10.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_10.TabIndex = 81
        Me._btnDatosTot_10.Text = "Mensaje Fiscal"
        Me._btnDatosTot_10.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_9
        '
        Me._btnDatosTot_9.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_9.Location = New System.Drawing.Point(1, 207)
        Me._btnDatosTot_9.Name = "_btnDatosTot_9"
        Me._btnDatosTot_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_9.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_9.TabIndex = 79
        Me._btnDatosTot_9.Text = "Mensaje Normal"
        Me._btnDatosTot_9.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_8
        '
        Me._btnDatosTot_8.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_8.Location = New System.Drawing.Point(1, 187)
        Me._btnDatosTot_8.Name = "_btnDatosTot_8"
        Me._btnDatosTot_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_8.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_8.TabIndex = 78
        Me._btnDatosTot_8.Text = "Cantidad Letra"
        Me._btnDatosTot_8.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_7
        '
        Me._btnDatosTot_7.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_7.Location = New System.Drawing.Point(1, 167)
        Me._btnDatosTot_7.Name = "_btnDatosTot_7"
        Me._btnDatosTot_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_7.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_7.TabIndex = 77
        Me._btnDatosTot_7.Text = "Cambio"
        Me._btnDatosTot_7.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_6
        '
        Me._btnDatosTot_6.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_6.Location = New System.Drawing.Point(1, 147)
        Me._btnDatosTot_6.Name = "_btnDatosTot_6"
        Me._btnDatosTot_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_6.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_6.TabIndex = 76
        Me._btnDatosTot_6.Text = "Pago"
        Me._btnDatosTot_6.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_5
        '
        Me._btnDatosTot_5.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_5.Location = New System.Drawing.Point(1, 127)
        Me._btnDatosTot_5.Name = "_btnDatosTot_5"
        Me._btnDatosTot_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_5.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_5.TabIndex = 75
        Me._btnDatosTot_5.Text = "Total Pesos"
        Me._btnDatosTot_5.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_4
        '
        Me._btnDatosTot_4.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_4.Location = New System.Drawing.Point(1, 107)
        Me._btnDatosTot_4.Name = "_btnDatosTot_4"
        Me._btnDatosTot_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_4.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_4.TabIndex = 74
        Me._btnDatosTot_4.Text = "Formas de Pago"
        Me._btnDatosTot_4.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_3
        '
        Me._btnDatosTot_3.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_3.Location = New System.Drawing.Point(1, 87)
        Me._btnDatosTot_3.Name = "_btnDatosTot_3"
        Me._btnDatosTot_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_3.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_3.TabIndex = 73
        Me._btnDatosTot_3.Text = "Redondeo"
        Me._btnDatosTot_3.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_2
        '
        Me._btnDatosTot_2.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_2.Location = New System.Drawing.Point(1, 67)
        Me._btnDatosTot_2.Name = "_btnDatosTot_2"
        Me._btnDatosTot_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_2.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_2.TabIndex = 72
        Me._btnDatosTot_2.Text = "I.V.A."
        Me._btnDatosTot_2.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_35
        '
        Me._btnDatosTot_35.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_35.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_35.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_35.Location = New System.Drawing.Point(1, 47)
        Me._btnDatosTot_35.Name = "_btnDatosTot_35"
        Me._btnDatosTot_35.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_35.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_35.TabIndex = 86
        Me._btnDatosTot_35.Text = "Descuento sin IVA"
        Me._btnDatosTot_35.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_1
        '
        Me._btnDatosTot_1.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_1.Location = New System.Drawing.Point(1, 27)
        Me._btnDatosTot_1.Name = "_btnDatosTot_1"
        Me._btnDatosTot_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_1.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_1.TabIndex = 71
        Me._btnDatosTot_1.Text = "Descuento con IVA"
        Me._btnDatosTot_1.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_0
        '
        Me._btnDatosTot_0.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_0.Location = New System.Drawing.Point(1, 7)
        Me._btnDatosTot_0.Name = "_btnDatosTot_0"
        Me._btnDatosTot_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_0.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_0.TabIndex = 70
        Me._btnDatosTot_0.Text = "SubTotal"
        Me._btnDatosTot_0.UseVisualStyleBackColor = False
        '
        'fraTotalesCredito
        '
        Me.fraTotalesCredito.BackColor = System.Drawing.SystemColors.Control
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_30)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_29)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_34)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_32)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_33)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_31)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_28)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_27)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_26)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_25)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_24)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_23)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_22)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_21)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_20)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_19)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_18)
        Me.fraTotalesCredito.Controls.Add(Me._btnDatosTot_17)
        Me.fraTotalesCredito.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTotalesCredito.Location = New System.Drawing.Point(648, 27)
        Me.fraTotalesCredito.Name = "fraTotalesCredito"
        Me.fraTotalesCredito.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTotalesCredito.Size = New System.Drawing.Size(135, 429)
        Me.fraTotalesCredito.TabIndex = 53
        Me.fraTotalesCredito.TabStop = False
        Me.fraTotalesCredito.Visible = False
        '
        '_btnDatosTot_30
        '
        Me._btnDatosTot_30.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_30.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_30.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_30.Location = New System.Drawing.Point(1, 267)
        Me._btnDatosTot_30.Name = "_btnDatosTot_30"
        Me._btnDatosTot_30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_30.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_30.TabIndex = 91
        Me._btnDatosTot_30.Text = "Moneda"
        Me._btnDatosTot_30.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_29
        '
        Me._btnDatosTot_29.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_29.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_29.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_29.Location = New System.Drawing.Point(1, 247)
        Me._btnDatosTot_29.Name = "_btnDatosTot_29"
        Me._btnDatosTot_29.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_29.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_29.TabIndex = 57
        Me._btnDatosTot_29.Text = "Total Piezas"
        Me._btnDatosTot_29.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_34
        '
        Me._btnDatosTot_34.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_34.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_34.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_34.Location = New System.Drawing.Point(67, 407)
        Me._btnDatosTot_34.Name = "_btnDatosTot_34"
        Me._btnDatosTot_34.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_34.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosTot_34.TabIndex = 59
        Me._btnDatosTot_34.Text = "Limpiar"
        Me._btnDatosTot_34.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_32
        '
        Me._btnDatosTot_32.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_32.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_32.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_32.Location = New System.Drawing.Point(67, 387)
        Me._btnDatosTot_32.Name = "_btnDatosTot_32"
        Me._btnDatosTot_32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_32.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosTot_32.TabIndex = 54
        Me._btnDatosTot_32.Text = "x"
        Me._btnDatosTot_32.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_33
        '
        Me._btnDatosTot_33.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_33.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_33.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_33.Location = New System.Drawing.Point(1, 407)
        Me._btnDatosTot_33.Name = "_btnDatosTot_33"
        Me._btnDatosTot_33.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_33.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosTot_33.TabIndex = 55
        Me._btnDatosTot_33.Text = "-"
        Me._btnDatosTot_33.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_31
        '
        Me._btnDatosTot_31.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_31.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_31.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_31.Location = New System.Drawing.Point(1, 387)
        Me._btnDatosTot_31.Name = "_btnDatosTot_31"
        Me._btnDatosTot_31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_31.Size = New System.Drawing.Size(67, 21)
        Me._btnDatosTot_31.TabIndex = 56
        Me._btnDatosTot_31.Text = "+"
        Me._btnDatosTot_31.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_28
        '
        Me._btnDatosTot_28.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_28.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_28.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_28.Location = New System.Drawing.Point(1, 227)
        Me._btnDatosTot_28.Name = "_btnDatosTot_28"
        Me._btnDatosTot_28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_28.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_28.TabIndex = 58
        Me._btnDatosTot_28.Text = "Mensaje Crédito"
        Me._btnDatosTot_28.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_27
        '
        Me._btnDatosTot_27.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_27.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_27.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_27.Location = New System.Drawing.Point(1, 207)
        Me._btnDatosTot_27.Name = "_btnDatosTot_27"
        Me._btnDatosTot_27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_27.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_27.TabIndex = 60
        Me._btnDatosTot_27.Text = "Mensaje Fiscal"
        Me._btnDatosTot_27.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_26
        '
        Me._btnDatosTot_26.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_26.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_26.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_26.Location = New System.Drawing.Point(1, 187)
        Me._btnDatosTot_26.Name = "_btnDatosTot_26"
        Me._btnDatosTot_26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_26.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_26.TabIndex = 61
        Me._btnDatosTot_26.Text = "Mensaje Normal"
        Me._btnDatosTot_26.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_25
        '
        Me._btnDatosTot_25.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_25.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_25.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_25.Location = New System.Drawing.Point(1, 167)
        Me._btnDatosTot_25.Name = "_btnDatosTot_25"
        Me._btnDatosTot_25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_25.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_25.TabIndex = 62
        Me._btnDatosTot_25.Text = "Cantidad Letra"
        Me._btnDatosTot_25.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_24
        '
        Me._btnDatosTot_24.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_24.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_24.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_24.Location = New System.Drawing.Point(1, 147)
        Me._btnDatosTot_24.Name = "_btnDatosTot_24"
        Me._btnDatosTot_24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_24.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_24.TabIndex = 63
        Me._btnDatosTot_24.Text = "Saldo"
        Me._btnDatosTot_24.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_23
        '
        Me._btnDatosTot_23.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_23.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_23.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_23.Location = New System.Drawing.Point(1, 127)
        Me._btnDatosTot_23.Name = "_btnDatosTot_23"
        Me._btnDatosTot_23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_23.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_23.TabIndex = 64
        Me._btnDatosTot_23.Text = "Anticipo"
        Me._btnDatosTot_23.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_22
        '
        Me._btnDatosTot_22.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_22.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_22.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_22.Location = New System.Drawing.Point(1, 107)
        Me._btnDatosTot_22.Name = "_btnDatosTot_22"
        Me._btnDatosTot_22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_22.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_22.TabIndex = 65
        Me._btnDatosTot_22.Text = "Total Pesos"
        Me._btnDatosTot_22.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_21
        '
        Me._btnDatosTot_21.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_21.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_21.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_21.Location = New System.Drawing.Point(1, 87)
        Me._btnDatosTot_21.Name = "_btnDatosTot_21"
        Me._btnDatosTot_21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_21.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_21.TabIndex = 66
        Me._btnDatosTot_21.Text = "Redondeo"
        Me._btnDatosTot_21.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_20
        '
        Me._btnDatosTot_20.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_20.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_20.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_20.Location = New System.Drawing.Point(1, 67)
        Me._btnDatosTot_20.Name = "_btnDatosTot_20"
        Me._btnDatosTot_20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_20.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_20.TabIndex = 67
        Me._btnDatosTot_20.Text = "I.V.A."
        Me._btnDatosTot_20.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_19
        '
        Me._btnDatosTot_19.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_19.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_19.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_19.Location = New System.Drawing.Point(1, 47)
        Me._btnDatosTot_19.Name = "_btnDatosTot_19"
        Me._btnDatosTot_19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_19.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_19.TabIndex = 87
        Me._btnDatosTot_19.Text = "Descuento sin IVA"
        Me._btnDatosTot_19.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_18
        '
        Me._btnDatosTot_18.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_18.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_18.Location = New System.Drawing.Point(1, 27)
        Me._btnDatosTot_18.Name = "_btnDatosTot_18"
        Me._btnDatosTot_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_18.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_18.TabIndex = 88
        Me._btnDatosTot_18.Text = "Descuento con IVA"
        Me._btnDatosTot_18.UseVisualStyleBackColor = False
        '
        '_btnDatosTot_17
        '
        Me._btnDatosTot_17.BackColor = System.Drawing.SystemColors.Control
        Me._btnDatosTot_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._btnDatosTot_17.ForeColor = System.Drawing.SystemColors.ControlText
        Me._btnDatosTot_17.Location = New System.Drawing.Point(1, 7)
        Me._btnDatosTot_17.Name = "_btnDatosTot_17"
        Me._btnDatosTot_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._btnDatosTot_17.Size = New System.Drawing.Size(134, 21)
        Me._btnDatosTot_17.TabIndex = 68
        Me._btnDatosTot_17.Text = "SubTotal"
        Me._btnDatosTot_17.UseVisualStyleBackColor = False
        '
        '_Marco_4
        '
        Me._Marco_4.BackColor = System.Drawing.SystemColors.Control
        Me._Marco_4.Controls.Add(Me._FlexDetalle_2)
        Me._Marco_4.Controls.Add(Me._lblFormula_2)
        Me._Marco_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Marco_4.Location = New System.Drawing.Point(9, 27)
        Me._Marco_4.Name = "_Marco_4"
        Me._Marco_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Marco_4.Size = New System.Drawing.Size(630, 429)
        Me._Marco_4.TabIndex = 46
        Me._Marco_4.TabStop = False
        '
        '_FlexDetalle_2
        '
        Me._FlexDetalle_2.DataSource = Nothing
        Me._FlexDetalle_2.Location = New System.Drawing.Point(9, 71)
        Me._FlexDetalle_2.Name = "_FlexDetalle_2"
        Me._FlexDetalle_2.OcxState = CType(resources.GetObject("_FlexDetalle_2.OcxState"), System.Windows.Forms.AxHost.State)
        Me._FlexDetalle_2.Size = New System.Drawing.Size(612, 346)
        Me._FlexDetalle_2.TabIndex = 9
        '
        '_lblFormula_2
        '
        Me._lblFormula_2.BackColor = System.Drawing.Color.FromArgb(CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer))
        Me._lblFormula_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._lblFormula_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblFormula_2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._lblFormula_2.Location = New System.Drawing.Point(9, 43)
        Me._lblFormula_2.Name = "_lblFormula_2"
        Me._lblFormula_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblFormula_2.Size = New System.Drawing.Size(613, 22)
        Me._lblFormula_2.TabIndex = 8
        Me._lblFormula_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'dbcSucursales
        '
        Me.dbcSucursales.Location = New System.Drawing.Point(80, 24)
        Me.dbcSucursales.Name = "dbcSucursales"
        Me.dbcSucursales.Size = New System.Drawing.Size(155, 21)
        Me.dbcSucursales.TabIndex = 1
        '
        '_Label2_1
        '
        Me._Label2_1.BackColor = System.Drawing.SystemColors.Info
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._Label2_1.Location = New System.Drawing.Point(440, 550)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(137, 16)
        Me._Label2_1.TabIndex = 51
        Me._Label2_1.Text = "Insert: Inserta Renglón"
        '
        '_Label2_0
        '
        Me._Label2_0.BackColor = System.Drawing.SystemColors.Info
        Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_0.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer))
        Me._Label2_0.Location = New System.Drawing.Point(192, 550)
        Me._Label2_0.Name = "_Label2_0"
        Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_0.Size = New System.Drawing.Size(137, 16)
        Me._Label2_0.TabIndex = 50
        Me._Label2_0.Text = "Supr: Elimina Renglón"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Info
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Label3.Location = New System.Drawing.Point(8, 547)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(785, 21)
        Me.Label3.TabIndex = 52
        '
        'FlexDetalle
        '
        '
        'btnDatosDet
        '
        '
        'btnDatosEnc
        '
        '
        'btnDatosTot
        '
        '
        'frmPVConfigTicketVenta
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ClientSize = New System.Drawing.Size(808, 573)
        Me.Controls.Add(Me.chkAplicarSucursales)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.TxtSaltos)
        Me.Controls.Add(Me.TxtEtiqueta)
        Me.Controls.Add(Me.SSTab1)
        Me.Controls.Add(Me.dbcSucursales)
        Me.Controls.Add(Me._Label1_0)
        Me.Controls.Add(Me._Label2_1)
        Me.Controls.Add(Me._Label2_0)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(124, 118)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPVConfigTicketVenta"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Configuración del Ticket"
        Me.Frame1.ResumeLayout(False)
        Me.SSTab1.ResumeLayout(False)
        Me._SSTab1_TabPage0.ResumeLayout(False)
        Me._Marco_1.ResumeLayout(False)
        CType(Me._FlexDetalle_0, System.ComponentModel.ISupportInitialize).EndInit()
        Me._Marco_0.ResumeLayout(False)
        Me._SSTab1_TabPage1.ResumeLayout(False)
        Me._Marco_3.ResumeLayout(False)
        Me._Marco_2.ResumeLayout(False)
        CType(Me._FlexDetalle_1, System.ComponentModel.ISupportInitialize).EndInit()
        Me._SSTab1_TabPage2.ResumeLayout(False)
        Me.fraTotalesContado.ResumeLayout(False)
        Me.fraTotalesCredito.ResumeLayout(False)
        Me._Marco_4.ResumeLayout(False)
        CType(Me._FlexDetalle_2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FlexDetalle, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Marco, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDatosDet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDatosEnc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDatosTot, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblFormula, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

End Class