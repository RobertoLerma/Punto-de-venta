'**********************************************************************************************************************'
'*PROGRAMA: MODULO DE COMPARATIVO JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'


Option Strict Off
Option Explicit On
Imports Excel = Microsoft.Office.Interop.Excel
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility

Public Module ModComparativo

    Const C_COLDIAMES As Integer = 2
    Const C_COLMES As Integer = 2
    Const C_ROWHEADER As Integer = 7

    Public ObjExcel As Object 'Objeto Excel
    Public Libro As Excel.Workbook 'Libro del Objeto Excel
    Public Hoja As Excel.Worksheet 'Hoja del Objeto Libro

    Dim rsCompa As ADODB.Recordset
    Dim nColActual As Integer
    Dim Anterior As Byte
    Dim TotSucAnt As Object
    Dim TotSucAct As Decimal
    Dim TotAlMesAnt As Object
    Dim TotAlMesAct As Decimal

    Public Sub ContTotales()
        Dim Fecha As Date
        Fecha = CDate(VB.Day(Today) & "/" & Month(Today) - 1 & "/" & Year(Today))
        Fecha = CDate(Fecha)
        With Hoja
            .Range("A22:B22").Merge()
            .Range("A22:B22")._Default = "Total"
            .Range("A22:B22").Font.Bold = True
            .Range("A23:B23").Merge()
            .Range("A23:B23")._Default = "Al mes de " & ModEstandar.MesLetra(Fecha, True)
            .Range("A23:B23").Font.Bold = True
        End With
    End Sub

    Public Sub LlenaDatos(ByRef rsDatosAnt As ADODB.Recordset, ByRef rsDatosNue As ADODB.Recordset, ByRef NoSuclsAnt As Integer, ByRef NoSuclsNue As Integer, ByRef TipoReporte As Boolean, ByRef blnSucAnt As Boolean, ByRef blnSucNue As Boolean)
        'On Error GoTo Merr
        Dim I As Integer
        Dim nCodSucursal As Integer
        PasarExel(TipoReporte)
        If blnSucAnt Then
            'Crear instancia de Excel
            Anterior = 1

            'Pasar los datos anteriores
            nColActual = 0
            rsCompa = rsDatosAnt
            rsCompa.MoveFirst()
            For I = 1 To NoSuclsAnt
                nCodSucursal = rsCompa.Fields("CodSucursal").Value
                ContenedorSucl(I, Trim(rsCompa.Fields("DescSucursal").Value.ToString()), TipoReporte)
                LlenaCantidadesSuc(I, nCodSucursal, TipoReporte)
                If I = 1 And Not TipoReporte Then
                    ContTotales()
                End If
                SumaTot()
            Next I
            'Preparar el formato para los totales
            rsCompa.MoveFirst()
            ContenedorTotales(TipoReporte)
            LlenaCantidadesTot(rsCompa.Fields("CodSucursal").Value.ToString(), TipoReporte)
            rsCompa.MoveFirst()
            ContenedorVtaDiaria(TipoReporte)
            LlenaCantidadesVta(rsCompa.Fields("CodSucursal").Value.ToString(), TipoReporte)
        End If
        If blnSucNue Then
            If Not blnSucAnt Then
                Anterior = 1
                nColActual = 0
            Else
                Anterior = 0
            End If
            'Pasar los datos nuevos
            rsCompa = rsDatosNue
            rsCompa.MoveFirst()
            For I = 1 To NoSuclsNue
                nCodSucursal = rsCompa.Fields("CodSucursal").Value.ToString()
                ContenedorSucl(I, Trim(rsCompa.Fields("DescSucursal").Value.ToString()), TipoReporte)
                LlenaCantidadesSuc(I, nCodSucursal, TipoReporte)
                If Anterior = 1 Then
                    If I = 1 And Not TipoReporte Then
                        ContTotales()
                    End If
                End If
                SumaTot()
            Next I
            'Preparar el formato para los totales
            rsCompa.MoveFirst()
            ContenedorTotales(TipoReporte)
            LlenaCantidadesTot(rsCompa.Fields("CodSucursal").Value.ToString(), TipoReporte)
            rsCompa.MoveFirst()
            ContenedorVtaDiaria(TipoReporte)
            LlenaCantidadesVta(rsCompa.Fields("CodSucursal").Value.ToString(), TipoReporte)
        End If

        'Poner títulos del reporte Centrados y las Etiquetas del reporte
        TituloEtiquetas(TipoReporte)

        Dim fechaGuardado As String = AgregarHoraAFecha(Today)

        'Guardarlo
        'Si existe el archivo de hoy, borrarlo para no sobreescribirlo
        If TipoReporte Then
            If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\CM" & VB6.Format(fechaGuardado, "dd/MM/yyyy") & ".xls") <> "" Then
                Kill((gstrCorpoDriveLocal & "\Sistema\Informes\CM" & VB6.Format(fechaGuardado, "dd/MM/yyyy") & ".xls"))
            End If
            Libro.Sheets(1).Select()

            'Dim ruta As String = (gstrCorpoDriveLocal & "\Sistema\Informes\CM" & VB6.Format(fechaGuardado, "dd/MM/yyyy") & ".xls")
            Dim ruta As Object = My.Application.Info.DirectoryPath & "\Sistema\Informes\CMPrueba.xls"

            Libro.SaveAs(ruta, Excel.XlFileFormat.xlWorkbookNormal, "", "", False, False)
            'Libro.SaveAs(ruta, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            If MsgBox("Se ha creado el archivo de Excel " & ruta & ", ¿desea abrirlo ahora?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa) = MsgBoxResult.Yes Then
                ObjExcel.Visible = True
            Else
                CierraInstanciaExcel()
            End If
        Else
            If Dir(gstrCorpoDriveLocal & "\Sistema\Informes\CA" & VB6.Format(fechaGuardado, "dd/MM/yyyy") & ".xls") <> "" Then
                Kill((gstrCorpoDriveLocal & "\Sistema\Informes\CA" & VB6.Format(fechaGuardado, "dd/MM/yyyy") & ".xls"))
            End If
            Libro.Sheets(1).Select()
            Libro.SaveAs(Filename:=gstrCorpoDriveLocal & "\Sistema\Informes\CA" & VB6.Format(fechaGuardado, "dd/MM/yyyy") & ".xls", FileFormat:=Excel.XlWindowState.xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False)
            If MsgBox("Se ha creado el archivo de Excel " & gstrCorpoDriveLocal & "\Sistema\Informes\CA" & VB6.Format(fechaGuardado, "dd/MM/yyyy") & ".xls, ¿desea abrirlo ahora?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, gstrNombCortoEmpresa) = MsgBoxResult.Yes Then
                ObjExcel.Visible = True
            Else
                CierraInstanciaExcel()
            End If
        End If
        'Hacer visible el objeto
        'System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        'Merr:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub CierraInstanciaExcel()
        On Error Resume Next
        'Cierro el libro
        Libro.Close()
        'Cierro la aplicacion
        ObjExcel.Quit()
        'Libero la memoria de mis variables
        ObjExcel = Nothing
        Libro = Nothing
        Hoja = Nothing
    End Sub

    Private Sub TituloEtiquetas(ByRef TipoReporte As Boolean)
        Dim msgMoneda As String
        Dim msgIVA As String
        nColActual = nColActual + 3
        If TipoReporte Then
            With Hoja
                'Título de la Empresa
                .Range(.Cells._Default(2, 2), .Cells._Default(2, 2)).Select()
                .Range(.Cells._Default(2, 2), .Cells._Default(2, 2))._Default = Trim(gstrNombCortoEmpresa)
                .Range(.Cells._Default(2, 2), .Cells._Default(2, 2)).Font.Bold = True
                .Range(.Cells._Default(2, 2), .Cells._Default(2, 2)).Font.Size = 12
                .Range(.Cells._Default(2, 2), .Cells._Default(2, nColActual)).Select()
                With .Range(.Cells._Default(2, 2), .Cells._Default(2, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(2, 2), .Cells._Default(2, nColActual)).Merge()
                'Título del Reporte
                .Range(.Cells._Default(3, 2), .Cells._Default(3, 2)).Select()
                .Range(.Cells._Default(3, 2), .Cells._Default(3, 2))._Default = "Comparativo de ventas diarias con año anterior"
                .Range(.Cells._Default(3, 2), .Cells._Default(3, 2)).Font.Bold = True
                .Range(.Cells._Default(3, 2), .Cells._Default(3, 2)).Font.Size = 10
                .Range(.Cells._Default(3, 2), .Cells._Default(3, nColActual)).Select()
                With .Range(.Cells._Default(3, 2), .Cells._Default(3, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(3, 2), .Cells._Default(3, nColActual)).Merge()
                'Título del Reporte
                .Range(.Cells._Default(4, 2), .Cells._Default(4, 2)).Select()
                '.Range(.Cells._Default(4, 2), .Cells._Default(4, 2))._Default = "Mes de " & Trim(frmVtasRPTVentasSalidadeMercanciaCompara.cboMes.Text) & " de " & Trim(frmVtasRPTVentasSalidadeMercanciaCompara.txtAnio.Text)
                .Range(.Cells._Default(4, 2), .Cells._Default(4, 2)).Font.Bold = True
                .Range(.Cells._Default(4, 2), .Cells._Default(4, 2)).Font.Size = 8
                .Range(.Cells._Default(4, 2), .Cells._Default(4, nColActual)).Select()
                With .Range(.Cells._Default(4, 2), .Cells._Default(4, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(4, 2), .Cells._Default(4, nColActual)).Merge()
                msgMoneda = ""
                msgIVA = ""
                'With frmVtasRPTVentasSalidadeMercanciaCompara
                '    Select Case True
                '        Case .optMoneda(0).Checked 'Cantidades en dólares
                '            msgMoneda = "** Las cantidades están expresadas en Dólares"
                '        Case Else
                '            msgMoneda = "** Las cantidades están expresadas en Pesos"
                '    End Select
                '    Select Case True
                '        Case .chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked
                '            msgIVA = "*** Las cantidades incluyen IVA"
                '        Case Else
                '            msgIVA = "*** Las cantidades NO incluyen IVA"
                '    End Select
                'End With
                .Range(.Cells._Default(5, 2), .Cells._Default(5, 2)).Select()
                .Range(.Cells._Default(5, 2), .Cells._Default(5, 2))._Default = msgMoneda
                .Range(.Cells._Default(5, 2), .Cells._Default(5, 2)).Font.Bold = False
                .Range(.Cells._Default(5, 2), .Cells._Default(5, 2)).Font.Size = 8
                .Range(.Cells._Default(5, 2), .Cells._Default(5, nColActual)).Select()
                With .Range(.Cells._Default(5, 2), .Cells._Default(5, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(5, 2), .Cells._Default(5, nColActual)).Merge()
                .Range(.Cells._Default(6, 2), .Cells._Default(6, 2)).Select()
                .Range(.Cells._Default(6, 2), .Cells._Default(6, 2))._Default = msgIVA
                .Range(.Cells._Default(6, 2), .Cells._Default(6, 2)).Font.Bold = False
                .Range(.Cells._Default(6, 2), .Cells._Default(6, 2)).Font.Size = 8
                .Range(.Cells._Default(6, 2), .Cells._Default(6, nColActual)).Select()
                With .Range(.Cells._Default(6, 2), .Cells._Default(6, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(6, 2), .Cells._Default(6, nColActual)).Merge()
                .Range(.Cells._Default(1, 1), .Cells._Default(1, 1)).Select()
            End With
        Else
            With Hoja
                'Título de la Empresa
                .Range(.Cells._Default(2, 2), .Cells._Default(2, 2)).Select()
                .Range(.Cells._Default(2, 2), .Cells._Default(2, 2))._Default = Trim(gstrNombCortoEmpresa)
                .Range(.Cells._Default(2, 2), .Cells._Default(2, 2)).Font.Bold = True
                .Range(.Cells._Default(2, 2), .Cells._Default(2, 2)).Font.Size = 12
                .Range(.Cells._Default(2, 2), .Cells._Default(2, nColActual)).Select()
                With .Range(.Cells._Default(2, 2), .Cells._Default(2, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(2, 2), .Cells._Default(2, nColActual)).Merge()
                'Título del Reporte
                .Range(.Cells._Default(3, 2), .Cells._Default(3, 2)).Select()
                .Range(.Cells._Default(3, 2), .Cells._Default(3, 2))._Default = "Comparativo de ventas anuales con año anterior"
                .Range(.Cells._Default(3, 2), .Cells._Default(3, 2)).Font.Bold = True
                .Range(.Cells._Default(3, 2), .Cells._Default(3, 2)).Font.Size = 10
                .Range(.Cells._Default(3, 2), .Cells._Default(3, nColActual)).Select()
                With .Range(.Cells._Default(3, 2), .Cells._Default(3, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(3, 2), .Cells._Default(3, nColActual)).Merge()
                'Título del Reporte
                .Range(.Cells._Default(4, 2), .Cells._Default(4, 2)).Select()
                '.Range(.Cells._Default(4, 2), .Cells._Default(4, 2))._Default = "Año de " & Trim(frmVtasRPTVentasSalidadeMercanciaCompara.txtAnio.Text)
                .Range(.Cells._Default(4, 2), .Cells._Default(4, 2)).Font.Bold = True
                .Range(.Cells._Default(4, 2), .Cells._Default(4, 2)).Font.Size = 8
                .Range(.Cells._Default(4, 2), .Cells._Default(4, nColActual)).Select()
                With .Range(.Cells._Default(4, 2), .Cells._Default(4, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(4, 2), .Cells._Default(4, nColActual)).Merge()
                msgMoneda = ""
                msgIVA = ""
                'With frmVtasRPTVentasSalidadeMercanciaCompara
                '    Select Case True
                '        Case .optMoneda(0).Checked 'Cantidades en dólares
                '            msgMoneda = "** Las cantidades están expresadas en Dólares"
                '        Case Else
                '            msgMoneda = "** Las cantidades están expresadas en Pesos"
                '    End Select
                '    Select Case True
                '        Case .chkImpuesto.CheckState = System.Windows.Forms.CheckState.Checked
                '            msgIVA = "*** Las cantidades incluyen IVA"
                '        Case Else
                '            msgIVA = "*** Las cantidades NO incluyen IVA"
                '    End Select
                'End With
                .Range(.Cells._Default(5, 2), .Cells._Default(5, 2)).Select()
                .Range(.Cells._Default(5, 2), .Cells._Default(5, 2))._Default = msgMoneda
                .Range(.Cells._Default(5, 2), .Cells._Default(5, 2)).Font.Bold = False
                .Range(.Cells._Default(5, 2), .Cells._Default(5, 2)).Font.Size = 8
                .Range(.Cells._Default(5, 2), .Cells._Default(5, nColActual)).Select()
                With .Range(.Cells._Default(5, 2), .Cells._Default(5, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(5, 2), .Cells._Default(5, nColActual)).Merge()
                .Range(.Cells._Default(6, 2), .Cells._Default(6, 2)).Select()
                .Range(.Cells._Default(6, 2), .Cells._Default(6, 2))._Default = msgIVA
                .Range(.Cells._Default(6, 2), .Cells._Default(6, 2)).Font.Bold = False
                .Range(.Cells._Default(6, 2), .Cells._Default(6, 2)).Font.Size = 8
                .Range(.Cells._Default(6, 2), .Cells._Default(6, nColActual)).Select()
                With .Range(.Cells._Default(6, 2), .Cells._Default(6, nColActual))
                    .HorizontalAlignment = Excel.Constants.xlLeft
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(6, 2), .Cells._Default(6, nColActual)).Merge()
                .Range(.Cells._Default(1, 1), .Cells._Default(1, 1)).Select()
            End With
        End If
    End Sub

    Sub SumaTot()
        Dim nContador As Integer
        nContador = C_ROWHEADER + 2
        With Hoja
            .Range(.Cells._Default(nContador + 13, nColActual), .Cells._Default(nContador + 13, nColActual)).Select()
            .Range(.Cells._Default(nContador + 13, nColActual), .Cells._Default(nContador + 13, nColActual)).NumberFormat = "#,##0.00_);[Black](#,##0.00)"
            .Range(.Cells._Default(nContador + 13, nColActual), .Cells._Default(nContador + 13, nColActual))._Default = TotSucAnt
            .Range(.Cells._Default(nContador + 13, nColActual + 1), .Cells._Default(nContador + 13, nColActual + 1)).Select()
            .Range(.Cells._Default(nContador + 13, nColActual + 1), .Cells._Default(nContador + 13, nColActual + 1)).NumberFormat = "#,##0.00_);[Black](#,##0.00)"
            .Range(.Cells._Default(nContador + 13, nColActual + 1), .Cells._Default(nContador + 13, nColActual + 1))._Default = TotSucAct
            .Range(.Cells._Default(nContador + 14, nColActual), .Cells._Default(nContador + 14, nColActual)).Select()
            .Range(.Cells._Default(nContador + 14, nColActual), .Cells._Default(nContador + 14, nColActual)).NumberFormat = "#,##0.00_);[Black](#,##0.00)"
            .Range(.Cells._Default(nContador + 14, nColActual), .Cells._Default(nContador + 14, nColActual))._Default = TotAlMesAnt
            .Range(.Cells._Default(nContador + 14, nColActual + 1), .Cells._Default(nContador + 14, nColActual + 1)).Select()
            .Range(.Cells._Default(nContador + 14, nColActual + 1), .Cells._Default(nContador + 14, nColActual + 1)).NumberFormat = "#,##0.00_);[Black](#,##0.00)"
            .Range(.Cells._Default(nContador + 14, nColActual + 1), .Cells._Default(nContador + 14, nColActual + 1))._Default = TotAlMesAct
        End With
    End Sub

    Private Sub LlenaCantidadesSuc(ByRef nSucl As Integer, ByRef nCodSucursal As Integer, ByRef TipoReporte As Boolean)
        Dim nContador As Integer
        Dim nSucursalActual As Integer
        nContador = C_ROWHEADER + 2
        nSucursalActual = rsCompa.Fields("CodSucursal").Value
        TotSucAct = 0
        TotSucAnt = 0
        TotAlMesAct = 0
        TotAlMesAnt = 0
        If TipoReporte Then
            While nCodSucursal = nSucursalActual
                With Hoja
                    nContador = nContador + 1
                    If Anterior = 1 Then
                        If nSucl = 1 Then
                            'Llenar los días de mes
                            .Range(.Cells._Default(nContador, nColActual - 1), .Cells._Default(nContador, nColActual - 1)).Select()
                            .Range(.Cells._Default(nContador, nColActual - 1), .Cells._Default(nContador, nColActual - 1))._Default = rsCompa.Fields("Dia").Value
                        End If
                    End If
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual)).Select()
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual))._Default = rsCompa.Fields("SaldoDiarioAn").Value
                    .Range(.Cells._Default(nContador, nColActual + 1), .Cells._Default(nContador, nColActual + 1))._Default = rsCompa.Fields("AcumAnterior").Value
                    .Range(.Cells._Default(nContador, nColActual + 2), .Cells._Default(nContador, nColActual + 2))._Default = rsCompa.Fields("SaldoDiarioAc").Value
                    .Range(.Cells._Default(nContador, nColActual + 3), .Cells._Default(nContador, nColActual + 3))._Default = rsCompa.Fields("AcumActual").Value
                End With
                rsCompa.MoveNext()
                If rsCompa.EOF Then
                    nCodSucursal = 0
                Else
                    nSucursalActual = rsCompa.Fields("CodSucursal").Value
                End If
            End While
        Else
            While nCodSucursal = nSucursalActual
                With Hoja
                    nContador = nContador + 1
                    If Anterior = 1 Then
                        If nSucl = 1 Then
                            'Llenar los días de mes
                            .Range(.Cells._Default(nContador, nColActual - 1), .Cells._Default(nContador, nColActual - 1)).Select()
                            .Range(.Cells._Default(nContador, nColActual - 1), .Cells._Default(nContador, nColActual - 1))._Default = rsCompa.Fields("descMes").Value
                        End If
                    End If
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual)).Select()
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual))._Default = rsCompa.Fields("ntotalesp").Value
                    .Range(.Cells._Default(nContador, nColActual + 1), .Cells._Default(nContador, nColActual + 1))._Default = rsCompa.Fields("ntotalesa").Value
                    TotSucAct = TotSucAct + rsCompa.Fields("ntotalesa").Value
                    TotSucAnt = TotSucAnt + rsCompa.Fields("ntotalesp").Value
                    If rsCompa.Fields("Mes").Value < Month(Today) Then
                        TotAlMesAct = TotAlMesAct + rsCompa.Fields("ntotalesa").Value
                        TotAlMesAnt = TotAlMesAnt + rsCompa.Fields("ntotalesp").Value
                    End If
                    '.Range(.Cells(nContador, nColActual + 2), .Cells(nContador, nColActual + 2)) = rsCompa!SaldoDiarioAc
                    '.Range(.Cells(nContador, nColActual + 3), .Cells(nContador, nColActual + 3)) = rsCompa!AcumActual
                End With
                rsCompa.MoveNext()
                If rsCompa.EOF Then
                    nCodSucursal = 0
                Else
                    nSucursalActual = rsCompa.Fields("CodSucursal").Value
                End If
            End While
        End If
    End Sub

    Private Sub LlenaCantidadesTot(ByRef nCodSucursal As Integer, ByRef TipoReporte As Boolean)
        Dim nContador As Integer
        Dim nSucursalActual As Integer
        nContador = C_ROWHEADER + 2
        nSucursalActual = rsCompa.Fields("CodSucursal").Value
        If TipoReporte Then
            While nCodSucursal = nSucursalActual
                nContador = nContador + 1
                With Hoja
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual)).Select()
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual))._Default = rsCompa.Fields("ntotalesp").Value
                    .Range(.Cells._Default(nContador, nColActual + 1), .Cells._Default(nContador, nColActual + 1))._Default = rsCompa.Fields("ntotalesa").Value
                    .Range(.Cells._Default(nContador, nColActual + 2), .Cells._Default(nContador, nColActual + 2))._Default = rsCompa.Fields("NTotalesPorc").Value
                End With
                rsCompa.MoveNext()
                If rsCompa.EOF Then
                    nCodSucursal = 0
                Else
                    nSucursalActual = rsCompa.Fields("CodSucursal").Value
                End If
            End While
        Else
            While nCodSucursal = nSucursalActual
                nContador = nContador + 1
                With Hoja
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual)).Select()
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual))._Default = rsCompa.Fields("totalesp").Value
                    .Range(.Cells._Default(nContador, nColActual + 1), .Cells._Default(nContador, nColActual + 1))._Default = rsCompa.Fields("totalesa").Value
                    If rsCompa.Fields("totalesp").Value = 0 Then
                        .Range(.Cells._Default(nContador, nColActual + 2), .Cells._Default(nContador, nColActual + 2))._Default = 0
                    Else
                        .Range(.Cells._Default(nContador, nColActual + 2), .Cells._Default(nContador, nColActual + 2))._Default = (rsCompa.Fields("totalesa").Value - rsCompa.Fields("totalesp").Value) / rsCompa.Fields("totalesp").Value
                    End If
                End With
                rsCompa.MoveNext()
                If rsCompa.EOF Then
                    nCodSucursal = 0
                Else
                    nSucursalActual = rsCompa.Fields("CodSucursal").Value
                End If
            End While
        End If
    End Sub

    Private Sub LlenaCantidadesVta(ByRef nCodSucursal As Integer, ByRef TipoReporte As Boolean)
        Dim nContador As Integer
        Dim nSucursalActual As Integer
        nContador = C_ROWHEADER + 2
        nSucursalActual = rsCompa.Fields("CodSucursal").Value
        If TipoReporte Then
            While nCodSucursal = nSucursalActual
                nContador = nContador + 1
                With Hoja
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual)).Select()
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual))._Default = rsCompa.Fields("vtadiariap").Value
                    .Range(.Cells._Default(nContador, nColActual + 1), .Cells._Default(nContador, nColActual + 1))._Default = rsCompa.Fields("VtaDiariaa").Value
                    .Range(.Cells._Default(nContador, nColActual + 2), .Cells._Default(nContador, nColActual + 2))._Default = rsCompa.Fields("VtaDiariaDif").Value
                End With
                rsCompa.MoveNext()
                If rsCompa.EOF Then
                    nCodSucursal = 0
                Else
                    nSucursalActual = rsCompa.Fields("CodSucursal").Value
                End If
            End While
        Else
            While nCodSucursal = nSucursalActual
                nContador = nContador + 1
                With Hoja
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual)).Select()
                    .Range(.Cells._Default(nContador, nColActual), .Cells._Default(nContador, nColActual))._Default = rsCompa.Fields("promant").Value
                    .Range(.Cells._Default(nContador, nColActual + 1), .Cells._Default(nContador, nColActual + 1))._Default = rsCompa.Fields("promact").Value
                    .Range(.Cells._Default(nContador, nColActual + 2), .Cells._Default(nContador, nColActual + 2))._Default = rsCompa.Fields("dif").Value
                End With
                rsCompa.MoveNext()
                If rsCompa.EOF Then
                    nCodSucursal = 0
                Else
                    nSucursalActual = rsCompa.Fields("CodSucursal").Value
                End If
            End While
        End If
    End Sub

    Private Sub ContenedorSucl(ByRef nSucl As Integer, ByRef cDescSucursal As String, ByRef TipoReporte As Boolean)
        If TipoReporte Then
            If nSucl = 1 Then
                'Es la primer sucursal
                'Actualiza la variable de columna actual
                nColActual = nColActual + (C_COLDIAMES + nSucl)
            Else
                nColActual = nColActual + 4
            End If
            With Hoja
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual))._Default = cDescSucursal
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 3)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 3))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 3)).Merge()
                'Segunda fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual))._Default = "ANTERIOR"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2))._Default = "ACTUAL"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 1))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 1)).Merge()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 3)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 3))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 3)).Merge()
                'Tercer fila del encabezado alineada a la derecha
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual))._Default = "Norm."
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 1), .Cells._Default(C_ROWHEADER + 2, nColActual + 1))._Default = "Acum."
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 2), .Cells._Default(C_ROWHEADER + 2, nColActual + 2))._Default = "Norm."
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 3), .Cells._Default(C_ROWHEADER + 2, nColActual + 3))._Default = "Acum."
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3))
                    .HorizontalAlignment = Excel.Constants.xlRight
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                'Ponerle color a las celdas de encabezado (Color gris claro)
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Interior
                    .ColorIndex = 15
                    .Pattern = Excel.Constants.xlSolid
                End With
                'Poner el formato a las celdas que contendrán números
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 3)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 3)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                'Establecer el color que deben tener las columnas
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual)).Font.ColorIndex = 10 'Verde
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Font.ColorIndex = 41 'Azul
                'Poner los bordes que delimitarán las celdas
                'Bordes del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3))
                    .Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                    .Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 3)).Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 3)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 3)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 3)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 3)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 3)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 3)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 3)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
            End With
        Else
            If nSucl = 1 Then
                'Es la primer sucursal
                'Actualiza la variable de columna actual
                nColActual = nColActual + (C_COLDIAMES + nSucl)
            Else
                nColActual = nColActual + 2
            End If
            With Hoja
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual))._Default = cDescSucursal
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 1))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 1)).Merge()
                'Segunda fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual))._Default = "ANTERIOR"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1))._Default = "ACTUAL"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 1))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                '.Range(.Cells(C_ROWHEADER + 1, nColActual), .Cells(C_ROWHEADER + 1, nColActual + 1)).Merge
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 1))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                '.Range(.Cells(C_ROWHEADER + 1, nColActual + 1), .Cells(C_ROWHEADER + 1, nColActual + 1)).Merge
                'Tercer fila del encabezado alineada a la derecha
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual))._Default = "Normal"
                '.Range(.Cells(C_ROWHEADER + 2, nColActual + 1), .Cells(C_ROWHEADER + 2, nColActual + 1)) = "Acum."
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 1), .Cells._Default(C_ROWHEADER + 2, nColActual + 1))._Default = "Normal"
                '.Range(.Cells(C_ROWHEADER + 2, nColActual + 3), .Cells(C_ROWHEADER + 2, nColActual + 3)) = "Acum."
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                'Ponerle color a las celdas de encabezado (Color gris claro)
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Interior
                    .ColorIndex = 15
                    .Pattern = Excel.Constants.xlSolid
                End With
                'Poner el formato a las celdas que contendrán números
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                'Establecer el color que deben tener las columnas
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual)).Font.ColorIndex = 10 'Verde
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 1), .Cells._Default(21, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 1), .Cells._Default(21, nColActual + 1)).Font.ColorIndex = 41 'Azul
                'Poner los bordes que delimitarán las celdas
                'Bordes del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual))
                    .Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                    .Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
            End With
        End If
    End Sub

    Private Sub ContenedorTotales(ByRef TipoReporte As Boolean)
        On Error Resume Next
        If TipoReporte Then
            With Hoja
                nColActual = nColActual + 4
                'Primer fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual))._Default = "TOTALES"
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2)).Merge()
                'Segunda fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual))._Default = "ANTERIOR"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1))._Default = "ACTUAL"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2))._Default = "PORC."
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 2))
                    .HorizontalAlignment = Excel.Constants.xlRight
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                'Tercer fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual))._Default = "NORMAL"
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 2), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 2), .Cells._Default(C_ROWHEADER + 2, nColActual + 2))._Default = "(Act.-Ant.)/Ant."
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Merge()
                'Pintar de gris parte del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Interior
                    .ColorIndex = 15
                    .Pattern = Excel.Constants.xlSolid
                End With
                'Poner el formato a las celdas que contendrán números
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                'Poner el formato de las celdas que contendrán porcentajes
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).NumberFormat = "0.00%"
                'Establecer el color que deben tener las columnas
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual)).Font.ColorIndex = 10 'Verde
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 1)).Font.ColorIndex = 41 'Azul
                'Poner los bordes que delimitarán las celdas
                'Bordes del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2))
                    .Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                    .Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades en la parte izquierda
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades en la parte derecha
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
            End With
        Else
            With Hoja
                nColActual = nColActual + 2
                'Primer fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual))._Default = "TOTALES"
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2)).Merge()
                'Segunda fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual))._Default = "ANTERIOR"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1))._Default = "ACTUAL"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2))._Default = "PORC."
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 2))
                    .HorizontalAlignment = Excel.Constants.xlRight
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                'Tercer fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual))._Default = "NORMAL"
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 2), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 2), .Cells._Default(C_ROWHEADER + 2, nColActual + 2))._Default = "(Act.-Ant.)/Ant."
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Merge()
                'Pintar de gris parte del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Interior
                    .ColorIndex = 15
                    .Pattern = Excel.Constants.xlSolid
                End With
                'Poner el formato a las celdas que contendrán números
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                'Poner el formato de las celdas que contendrán porcentajes
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).NumberFormat = "0.00%"
                'Establecer el color que deben tener las columnas
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual)).Font.ColorIndex = 10 'Verde
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 1)).Font.ColorIndex = 41 'Azul
                'Poner los bordes que delimitarán las celdas
                'Bordes del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2))
                    .Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                    .Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades en la parte izquierda
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades en la parte derecha
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
            End With
        End If
    End Sub

    Private Sub ContenedorVtaDiaria(ByRef TipoReporte As Boolean)
        On Error Resume Next
        If TipoReporte Then
            With Hoja
                nColActual = nColActual + 3
                'Primer fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual))._Default = "VENTA DIARIA"
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2)).Merge()
                'Segunda fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual))._Default = "ANTERIOR"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1))._Default = "ACTUAL"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2))._Default = "DIFERENCIA"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 2))
                    .HorizontalAlignment = Excel.Constants.xlRight
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                'Tercer fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual))._Default = "NORMAL"
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 2), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 2), .Cells._Default(C_ROWHEADER + 2, nColActual + 2))._Default = ""
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Merge()
                'Pintar de gris parte del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Interior
                    .ColorIndex = 15
                    .Pattern = Excel.Constants.xlSolid
                End With
                'Poner el formato a las celdas que contendrán números
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 2)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                'Establecer el color que deben tener las columnas
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual)).Font.ColorIndex = 10 'Verde
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 1)).Font.ColorIndex = 41 'Azul
                'Poner los bordes que delimitarán las celdas
                'Bordes del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2))
                    .Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                    .Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades en la parte izquierda
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(40, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades en la parte derecha
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(40, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
            End With
        Else
            With Hoja
                nColActual = nColActual + 3
                'Primer fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual))._Default = "VENTA DIARIA"
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER, nColActual + 2)).Merge()
                'Segunda fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual))._Default = "ANTERIOR"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 1), .Cells._Default(C_ROWHEADER + 1, nColActual + 1))._Default = "ACTUAL"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual + 2), .Cells._Default(C_ROWHEADER + 1, nColActual + 2))._Default = "DIFERENCIA"
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 1, nColActual + 2))
                    .HorizontalAlignment = Excel.Constants.xlRight
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                'Tercer fila del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual))._Default = "NORMAL"
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 2), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual + 2), .Cells._Default(C_ROWHEADER + 2, nColActual + 2))._Default = ""
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1))
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range(.Cells._Default(C_ROWHEADER + 2, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 1)).Merge()
                'Pintar de gris parte del encabezado
                .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER + 1, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Interior
                    .ColorIndex = 15
                    .Pattern = Excel.Constants.xlSolid
                End With
                'Poner el formato a las celdas que contendrán números
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 2)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                'Establecer el color que deben tener las columnas
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual)).Font.ColorIndex = 10 'Verde
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 1)).Font.ColorIndex = 41 'Azul
                'Poner los bordes que delimitarán las celdas
                'Bordes del encabezado
                .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Select()
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2))
                    .Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                    .Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlInsideVertical)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER, nColActual), .Cells._Default(C_ROWHEADER + 2, nColActual + 2)).Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades en la parte izquierda
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual), .Cells._Default(21, nColActual + 1)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                'Bordes del Contenedor de cantidades en la parte derecha
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Select()
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range(.Cells._Default(C_ROWHEADER + 3, nColActual + 2), .Cells._Default(21, nColActual + 2)).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
            End With
        End If
    End Sub

    Private Sub PasarExel(ByRef TipoReporte As Boolean)
        On Error GoTo Merr
        Dim lnNumHojas As Integer
        Dim rsLocal As ADODB.Recordset

        lnNumHojas = 1

        CierraInstanciaExcel()

        'Checar si no existe la carpeta para crearla
        'If Dir(gstrCorpoDriveLocal & "\Sistema\Informes", FileAttribute.Directory) = "" Then
        '    MkDir(gstrCorpoDriveLocal & "\Sistema\Informes")
        'End If

        If Dir("C:\Users\Consultor Vitek\Downloads\CORPORATIVO Y JOYERIA\CODIGO ANGEL WHA\CorporativoV1\CorporativoV1\Sistema\Informes\Informes", FileAttribute.Directory) = "" Then
            MkDir("C:\Users\Consultor Vitek\Downloads\CORPORATIVO Y JOYERIA\CODIGO ANGEL WHA\CorporativoV1\CorporativoV1\Sistema\Informes\Informes")
        End If


        ObjExcel = CreateObject("Excel.Application")
        Libro = ObjExcel.Workbooks.Add

        Hoja = Libro.Worksheets.Add
        Libro.Sheets(lnNumHojas).Select()
        Hoja = Libro.ActiveSheet
        'System.Windows.Forms.Application.DoEvents()
        LlenarInformacion(TipoReporte)
        'MsgBox "Se ha generado el archivo: C:\Informes\Comparativo" & Format(Date, "yyyymmdd") & ".xls", vbOKOnly + vbInformation, gstrNombCortoEmpresa
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
        End If
    End Sub

    Private Sub LlenarInformacion(ByRef TipoReporte As Boolean)
        On Error GoTo Merr
        'Empieza el reporte
        With Hoja
            'Cambia el tipo de letra y el tamaño
            .Cells.Select()
            With .Cells.Font
                .Name = "Arial"
                .Size = 8
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone
            End With
            'Crea la columna que contendrá los números de día del mes la dimensiona a 4 (no sé qué medida maneja)
            .Columns._Default("B:B").ColumnWidth = 4
            'Da formato a la columna que contendrá los números de día del mes
            FormatoColumnaDiaMes(TipoReporte)
        End With
Merr:
        If Err.Number <> 0 Then
            ModEstandar.MostrarError()
            Return
        End If
    End Sub

    Private Sub FormatoColumnaDiaMes(ByRef TipoReporte As Boolean)
        On Error Resume Next
        If TipoReporte Then
            With Hoja
                .Range("B7:B9").Select()
                With .Range("B7:B9")
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = True
                End With
                .Range("B7:B9").Merge()
                .Range("B7:B9")._Default = "Día"
                .Range("B7:B9").Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range("B7:B9").Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range("B7:B9").Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B7:B9").Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B7:B9").Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B7:B9").Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range("B7:B9").Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.Constants.xlNone
                .Range("B10:B40").Select()
                With .Range("B10:B40")
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range("B10:B40").Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range("B10:B40").Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range("B10:B40").Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B10:B40").Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B10:B40").Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B10:B40").Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range("B10:B40").Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.Constants.xlNone
                .Range("B7:B9").Select()
            End With
        Else
            With Hoja
                .Range("B7:B9").Select()
                With .Range("B7:B9")
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = True
                End With
                .Range("B7:B9").Merge()
                .Range("B7:B9")._Default = "Mes"
                .Range("B7:B9").Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range("B7:B9").Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range("B7:B9").Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B7:B9").Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B7:B9").Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B7:B9").Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range("B7:B9").Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.Constants.xlNone
                .Range("B10:B21").Select()
                With .Range("B10:B21")
                    .HorizontalAlignment = Excel.Constants.xlCenter
                    .VerticalAlignment = Excel.Constants.xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                    .MergeCells = False
                End With
                .Range("B10:B21").Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
                .Range("B10:B21").Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
                With .Range("B10:B21").Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B10:B21").Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B10:B21").Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                With .Range("B10:B21").Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = Excel.XlBorderWeight.xlThin
                    .ColorIndex = Excel.Constants.xlAutomatic
                End With
                .Range("B10:B21").Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.Constants.xlNone
                .Range("B7:B9").Select()
            End With
        End If
    End Sub
End Module