'**********************************************************************************************************************'
'*PROGRAMA: MODULO DE COMBOS JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports System.Data.Common
Imports System.Data.SqlClient
Imports ADODB

Public Module ModDCombo
    Dim rsLocal As New ADODB.Recordset
    Dim i As Integer

    'Parámetros
    '   1.- Sql   : sentencia de sql que filtrará la información
    '   2.- Tecla : Valor de la tecla que se presionó
    '       Nota  : El orden de los campos en sql serán código (Pk) y descripción
    Sub DCChange(ByRef Sql As String, ByRef tecla As Integer, Optional ByRef Combo As System.Windows.Forms.ComboBox = Nothing)
        'On Error GoTo Errores
        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.AppStarting
            If Combo Is Nothing Then 'Si no se le dio como parametro algun control se le accina el default
                'Combo = GetActiveControl()
            End If
            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
            rsLocal = Cmd.Execute

            'If tecla <> 0 Then
            '    If tecla <> 40 And tecla <> 38 And tecla <> 33 And tecla <> 34 And rsLocal.RecordCount > 0 Then
            '       ' Combo.Text = rsLocal
            '       ' Combo.Text = rsLocal.Fields(1).Name
            '    End If
            'End If

            'tecla = 0
            'Combo.Refresh()

            If (Combo.Name = "dbcBanco") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodBanco"
                Combo.DisplayMember = "DescBanco"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("DescBanco").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcDescMarca") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codMarca"
                Combo.DisplayMember = "descMarca"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descMarca").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcGrupo") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codGrupo"
                Combo.DisplayMember = "descGrupo"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descGrupo").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcDescFamilia") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcDescLinea") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codLinea"
                Combo.DisplayMember = "descLinea"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descLinea").ToString()
                '  'End With
                ''Next i
            End If



            If (Combo.Name = "dbcSucursal") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codAlmacen"
                Combo.DisplayMember = "descAlmacen"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descAlmacen").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcProveedor") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codProvAcreed"
                Combo.DisplayMember = "descProvAcreed"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descProvAcreed").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcProveedores") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codProvAcreed"
                Combo.DisplayMember = "descProvAcreed"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descProvAcreed").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcJFamilia") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcJLinea") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codLinea"
                Combo.DisplayMember = "descLinea"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descLinea").ToString()
                '  'End With
                ''Next i
            End If



            If (Combo.Name = "dbcJSubLinea") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codSubLinea"
                Combo.DisplayMember = "descSubLinea"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descSubLinea").ToString()
                '  'End With
                ''Next i
            End If



            If (Combo.Name = "dbcRModelo") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codModelo"
                Combo.DisplayMember = "descModelo"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descModelo").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcVFamilia") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcVLinea") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codLinea"
                Combo.DisplayMember = "descLinea"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descLinea").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcRMarca") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codMarca"
                Combo.DisplayMember = "descMarca"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descMarca").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcMaterial") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codTipoMaterial"
                Combo.DisplayMember = "descTipoMaterial"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descTipoMaterial").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcVendedor") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codVendedor"
                Combo.DisplayMember = "descVendedor"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descVendedor").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcCliente") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodCliente"
                Combo.DisplayMember = "DescCliente"
                'Combo.SelectedIndex = 1
                'Combo.Text = dt.Rows(i)("DescCliente").ToString()
                'End With
                ''Next i
            End If

            If (Combo.Name = "dbcTaller") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codTaller"
                Combo.DisplayMember = "descTaller"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descTaller").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcCaja") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodCaja"
                Combo.DisplayMember = "NumCaja"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("NumCaja").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcTipoReparacioN") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodGrupo"
                Combo.DisplayMember = "DescGrupo"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("DescGrupo").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcFamilia") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "_dbcFamilia_0" Or Combo.Name = "_dbcFamilia_1") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "_dbcLinea_0" Or Combo.Name = "_dbcLinea_1") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codLinea"
                Combo.DisplayMember = "descLinea"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descLinea").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcSubLinea") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codSubLinea"
                Combo.DisplayMember = "descSubLinea"
                'Combo.SelectedIndex = 1 
                '       ' Combo.Text = dt.Rows(i)("descSubLinea").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcKilates") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codKilates"
                Combo.DisplayMember = "descKilates"
                'Combo.SelectedIndex = 1 
                '       ' Combo.Text = dt.Rows(i)("descKilates").ToString()
                '  'End With
                ''Next i
            End If



            If (Combo.Name = "_dbcMaterial_0" Or Combo.Name = "_dbcMaterial_1" Or Combo.Name = "_dbcMaterial_2") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codTipoMaterial"
                Combo.DisplayMember = "descTipoMaterial"
                'Combo.SelectedIndex = 1  
                '       ' Combo.Text = dt.Rows(i)("descTipoMaterial").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "_cboUnidad_0" Or Combo.Name = "_cboUnidad_1" Or Combo.Name = "_cboUnidad_2") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codUnidad"
                Combo.DisplayMember = "descUnidad"
                'Combo.SelectedIndex = 1  
                '       ' Combo.Text = dt.Rows(i)("descUnidad").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "_cboAlmacen_0" Or Combo.Name = "_cboAlmacen_1" Or Combo.Name = "_cboAlmacen_2") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codAlmacenOrigen"
                Combo.DisplayMember = "descAlmacenOrigen"
                'Combo.SelectedIndex = 1  
                '       ' Combo.Text = dt.Rows(i)("descAlmacenOrigen").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "_dbcProveedor_0" Or Combo.Name = "_dbcProveedor_1" Or Combo.Name = "_dbcProveedor_2") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codProvAcreed"
                Combo.DisplayMember = "descProvAcreed"
                'Combo.SelectedIndex = 1  
                '       ' Combo.Text = dt.Rows(i)("descProvAcreed").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcOrigen") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodAlmacenOrigen"
                Combo.DisplayMember = "DescAlmacen"
                'Combo.SelectedIndex = 1   
                '       ' Combo.Text = dt.Rows(i)("DescAlmacen").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcMarca") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codMarca"
                Combo.DisplayMember = "descMarca"
                'Combo.SelectedIndex = 1   
                '       ' Combo.Text = dt.Rows(i)("descMarca").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcModelo") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codModelo"
                Combo.DisplayMember = "descModelo"
                'Combo.SelectedIndex = 1   
                '       ' Combo.Text = dt.Rows(i)("descModelo").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "_dbcGrupos_0" Or Combo.Name = "_dbcGrupos_1") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codUsuario"
                Combo.DisplayMember = "Nombre"
                'Combo.SelectedIndex = 1   
                ' Combo.Text = dt.Rows(i)("Nombre").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcModulo") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codModulo"
                Combo.DisplayMember = "descModulo"
                'Combo.SelectedIndex = 1   
                ' Combo.Text = dt.Rows(i)("descModulo").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcUsuarios") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codUsuario"
                Combo.DisplayMember = "Nombre"
                'Combo.SelectedIndex = 1   
                'Combo.Text = dt.Rows(i)("Nombre").ToString()
                'End With
                'Next i
            End If




            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
            'Errores:
        Catch ex As Exception
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            ModErrores.Errores()
        End Try
    End Sub

    'Parámetros
    '   1.- Sql   : sentencia de sql que filtrará la información
    Sub DCGotFocus(ByRef Sql As String, Optional ByRef Combo As System.Windows.Forms.ComboBox = Nothing)
        'On Error GoTo Errores
        Try
            '' Selecciono la información del control
            'If Combo Is Nothing Then 'Si no se le dio como parametro algun control se le accina el default
            '    Screen.ActiveForm.ActiveControl.SelStart = 0
            '    Screen.ActiveForm.ActiveControl.SelLength = Len(Screen.ActiveForm.ActiveControl.text)
            'Else
            '    Combo.SelectedIndex = 0
            '    Combo.SelectionLength = Len(Combo.Text)
            'End If

            ModEstandar.BorraCmd()
            Cmd.CommandText = "dbo.Up_Select_Datos"
            Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
            Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
            rsLocal = Cmd.Execute


            If (Combo.Name = "dbcBanco") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodBanco"
                Combo.DisplayMember = "DescBanco"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("DescBanco").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcDescMarca") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codMarca"
                Combo.DisplayMember = "descMarca"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descMarca").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcGrupo") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'for i = 0 to dt.rows.count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codGrupo"
                Combo.DisplayMember = "descGrupo"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descGrupo").ToString()
                '  'End With
                ''Next i
            End If

            If (Combo.Name = "dbcDescFamilia") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcDescLinea") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codLinea"
                Combo.DisplayMember = "descLinea"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descLinea").ToString()
                '  'End With
                ''Next i
            End If



            If (Combo.Name = "dbcSucursal" Or Combo.Name = "dbcSucursales" Or Combo.Name = "dbcSucOrigen" Or Combo.Name = "dbcSucDestino" Or Combo.Name = "dbcAlmacen") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                'If (Combo.Text > "") Then
                Combo.DataSource = dt
                Combo.ValueMember = "codAlmacen"
                Combo.DisplayMember = "descAlmacen"
                'End If

                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descAlmacen").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcProveedor" Or Combo.Name = "dbcProveedorAcreedor") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codProvAcreed"
                Combo.DisplayMember = "descProvAcreed"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descProvAcreed").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcProveedores") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codProvAcreed"
                Combo.DisplayMember = "descProvAcreed"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descProvAcreed").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcJFamilia") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                '  'End With
                ''Next i
            End If


            If (Combo.Name = "dbcJLinea") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                '   'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codLinea"
                Combo.DisplayMember = "descLinea"
                'Combo.SelectedIndex = 1
                '       ' Combo.Text = dt.Rows(i)("descLinea").ToString()
                '  'End With
                ''Next i
            End If



            If (Combo.Name = "dbcJSubLinea") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codSubLinea"
                Combo.DisplayMember = "descSubLinea"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descSubLinea").ToString()
                'End With
                'Next i
            End If



            If (Combo.Name = "dbcRModelo" Or Combo.Name = "dbcRmodelo") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codModelo"
                Combo.DisplayMember = "descModelo"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descModelo").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcVFamilia") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcVLinea") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codLinea"
                Combo.DisplayMember = "descLinea"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descLinea").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcRMarca" Or Combo.Name = "dbcRMarca") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codMarca"
                Combo.DisplayMember = "descMarca"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descMarca").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcMaterial") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codTipoMaterial"
                Combo.DisplayMember = "descTipoMaterial"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descTipoMaterial").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcVendedor") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codVendedor"
                Combo.DisplayMember = "descVendedor"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descVendedor").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcCliente") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodCliente"
                Combo.DisplayMember = "DescCliente"
                'Combo.SelectedIndex = 1
                'Combo.Text = dt.Rows(i)("DescCliente").ToString()
                'End With
                ''Next i
            End If


            If (Combo.Name = "dbcTaller") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codTaller"
                Combo.DisplayMember = "descTaller"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descTaller").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcCaja") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodCaja"
                Combo.DisplayMember = "NumCaja"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("NumCaja").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcTipoReparacion") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodGrupo"
                Combo.DisplayMember = "DescGrupo"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("DescGrupo").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcFamilia") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "_dbcFamilia_0" Or Combo.Name = "_dbcFamilia_1" Or Combo.Name = "dbcjFAmilia" Or Combo.Name = "dbcVFamilia") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codFamilia"
                Combo.DisplayMember = "descFamilia"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descFamilia").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "_dbcLinea_0" Or Combo.Name = "_dbcLinea_1" Or Combo.Name = "dbcJLinea" Or Combo.Name = "dbcVLinea") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codLinea"
                Combo.DisplayMember = "descLinea"
                'Combo.SelectedIndex = 1
                ' Combo.Text = dt.Rows(i)("descLinea").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcSubLinea" Or Combo.Name = "" Or Combo.Name = "dbcJSubLinea") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codSubLinea"
                Combo.DisplayMember = "descSubLinea"
                'Combo.SelectedIndex = 1 
                ' Combo.Text = dt.Rows(i)("descSubLinea").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcKilates") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codKilates"
                Combo.DisplayMember = "descKilates"
                'Combo.SelectedIndex = 1 
                ' Combo.Text = dt.Rows(i)("descKilates").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "_dbcMaterial_0" Or Combo.Name = "_dbcMaterial_1" Or Combo.Name = "_dbcMaterial_2") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codTipoMaterial"
                Combo.DisplayMember = "descTipoMaterial"
                'Combo.SelectedIndex = 1  
                ' Combo.Text = dt.Rows(i)("descTipoMaterial").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "_cboUnidad_0" Or Combo.Name = "_cboUnidad_1" Or Combo.Name = "_cboUnidad_2") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codUnidad"
                Combo.DisplayMember = "descUnidad"
                'Combo.SelectedIndex = 1  
                ' Combo.Text = dt.Rows(i)("descUnidad").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "_cboAlmacen_0" Or Combo.Name = "_cboAlmacen_1" Or Combo.Name = "_cboAlmacen_2") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codAlmacenOrigen"
                Combo.DisplayMember = "descAlmacenOrigen"
                'Combo.SelectedIndex = 1  
                ' Combo.Text = dt.Rows(i)("descAlmacenOrigen").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "_dbcProveedor_0" Or Combo.Name = "_dbcProveedor_1" Or Combo.Name = "_dbcProveedor_2") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codProvAcreed"
                Combo.DisplayMember = "descProvAcreed"
                'Combo.SelectedIndex = 1  
                ' Combo.Text = dt.Rows(i)("descProvAcreed").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcOrigen") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodAlmacenOrigen"
                Combo.DisplayMember = "DescAlmacen"
                'Combo.SelectedIndex = 1   
                ' Combo.Text = dt.Rows(i)("DescAlmacen").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcOrigen1") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodAlmacenOrigen"
                Combo.DisplayMember = "descAlmacenOrigen"
                'Combo.SelectedIndex = 1   
                ' Combo.Text = dt.Rows(i)("DescAlmacen").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcMarca") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codMarca"
                Combo.DisplayMember = "descMarca"
                'Combo.SelectedIndex = 1   
                ' Combo.Text = dt.Rows(i)("descMarca").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcModelo") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codModelo"
                Combo.DisplayMember = "descModelo"
                'Combo.SelectedIndex = 1   
                ' Combo.Text = dt.Rows(i)("descModelo").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "_dbcGrupos_0" Or Combo.Name = "_dbcGrupos_1") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codUsuario"
                Combo.DisplayMember = "Nombre"
                'Combo.SelectedIndex = 1   
                ' Combo.Text = dt.Rows(i)("Nombre").ToString()
                'End With
                'Next i
            End If


            If (Combo.Name = "dbcModulo") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codModulo"
                Combo.DisplayMember = "descModulo"
                'Combo.SelectedIndex = 1   
                ' Combo.Text = dt.Rows(i)("descModulo").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcUsuarios") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "codUsuario"
                Combo.DisplayMember = "Nombre"
                'Combo.SelectedIndex = 1   
                'Combo.Text = dt.Rows(i)("Nombre").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcCuentaBancaria") Then
                'Cmd.CommandTimeout = 1200
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                'For i = 0 To dt.Rows.Count - 1
                'With Combo
                Combo.DataSource = dt
                Combo.ValueMember = "CodBanco"
                Combo.DisplayMember = "CtaBancaria"
                'Combo.SelectedIndex = 1   
                'Combo.Text = dt.Rows(i)("Nombre").ToString()
                'End With
                'Next i
            End If

            If (Combo.Name = "dbcAgrupador") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                Combo.DataSource = dt
                Combo.ValueMember = "CodOrigenAplicR"
                Combo.DisplayMember = "DescOrigenAplicR"
            End If

            If (Combo.Name = "dbcRubro") Then
                Dim dt As New DataTable()
                Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                adapter.Fill(dt)
                Combo.DataSource = dt
                Combo.ValueMember = "CodRubro"
                Combo.DisplayMember = "DescRubro"
            End If




            'Errores:
        Catch ex As Exception
            If Err.Number <> 0 Then ModErrores.Errores()
        End Try
    End Sub

    'Parámetros
    '   1.- DataCombo   : Nombre del control DataCombo
    '   2.- Sql         : sentencia de sql que filtrará la información
    'Aqui si se pasa el control como parametro porke cuando se ejecuta este procedimiento ya el foco esta en otro control
    Sub DCLostFocus(ByRef DataCombo As System.Windows.Forms.ComboBox, ByRef Sql As String, ByRef nCodigo As Integer)
        'On Error GoTo Errores
        Try
            Dim RecCombo As ADODB.Recordset
            If DataCombo.Text <> "" Then
                ModEstandar.BorraCmd()
                Cmd.CommandText = "dbo.Up_Select_Datos"
                Cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Cmd.Parameters.Append(Cmd.CreateParameter("Renglon", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamReturnValue))
                '''Cmd.Parameters.Append Cmd.CreateParameter("Sentencia", adChar, adParamInput, 8000, gStrSql) - OJO - 03MAR2008 - MAVF
                Cmd.Parameters.Append(Cmd.CreateParameter("Sentencia", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, Sql))
                rsLocal = Cmd.Execute


                'If (DataCombo.Name = "dbcProveedores") Then
                '    Dim dt As New DataTable()
                '    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                '    adapter.Fill(dt)
                '    'For i = 0 To dt.Rows.Count - 1
                '       'With DataCombo
                '            DataCombo.DataSource = dt
                '            DataCombo.ValueMember = "codProvAcreed"
                '            DataCombo.DisplayMember = "descProvAcreed"
                '            'DataCombo.SelectedIndex = 1
                '           'DataCombo.Text = dt.Rows(i)("descProvAcreed").ToString()
                '      'End With
                '    'Next i
                'End If


                '    If rsLocal.RecordCount > 0 Then
                '       'DataCombo.Text = rsLocal.Fields(1).Name
                '        nCodigo = rsLocal.Fields(0).Value
                '       'DataCombo.Text = Trim(rsLocal.Fields(1).Value)
                '    Else
                '        nCodigo = 0
                '       'DataCombo.Text = ""
                '    End If
                'Else
                '    nCodigo = 0
                'End If

                'rsLocal = Nothing


                If (DataCombo.Name = "dbcBanco") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "CodBanco"
                    DataCombo.DisplayMember = "DescBanco"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("CodBanco").ToString()
                    'DataCombo.Text = dt.Rows(i)("DescBanco").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcDescMarca") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codMarca"
                    DataCombo.DisplayMember = "descMarca"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codMarca").ToString()
                    'DataCombo.Text = dt.Rows(i)("descMarca").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcGrupo") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codGrupo"
                    DataCombo.DisplayMember = "descGrupo"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codGrupo").ToString()
                    'DataCombo.Text = dt.Rows(i)("descGrupo").ToString()
                    'End With
                    'Next i
                End If

                If (DataCombo.Name = "dbcDescFamilia") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codFamilia"
                    DataCombo.DisplayMember = "descFamilia"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codFamilia").ToString()
                    'DataCombo.Text = dt.Rows(i)("descFamilia").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcDescLinea") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codLinea"
                    DataCombo.DisplayMember = "descLinea"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codLinea").ToString()
                    'DataCombo.Text = dt.Rows(i)("descLinea").ToString()
                    'End With
                    'Next i
                End If



                If (DataCombo.Name = "dbcSucursal" Or DataCombo.Name = "dbcSucursales" Or DataCombo.Name = "dbcSucOrigen" Or DataCombo.Name = "dbcAlmacen") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    'DataCombo.DataSource = dt
                    'DataCombo.ValueMember = "codAlmacen"
                    'DataCombo.DisplayMember = "descAlmacen"
                    'DataCombo.SelectedIndex = 1
                    If (dt.Rows.Count > 0) Then
                        nCodigo = dt.Rows(i)("codAlmacen").ToString()
                    End If
                    'DataCombo.Text = dt.Rows(i)("descAlmacen").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcProveedor" Or DataCombo.Name = "dbcProveedorAcreedor") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codProvAcreed"
                    DataCombo.DisplayMember = "descProvAcreed"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codProvAcreed").ToString()
                    'DataCombo.Text = dt.Rows(i)("descProvAcreed").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcProveedores") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codProvAcreed"
                    DataCombo.DisplayMember = "descProvAcreed"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codProvAcreed").ToString()
                    'DataCombo.Text = dt.Rows(i)("descProvAcreed").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcJFamilia") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codFamilia"
                    DataCombo.DisplayMember = "descFamilia"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codFamilia").ToString()
                    'DataCombo.Text = dt.Rows(i)("descFamilia").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcJLinea") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codLinea"
                    DataCombo.DisplayMember = "descLinea"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codLinea").ToString()
                    'DataCombo.Text = dt.Rows(i)("descLinea").ToString()
                    'End With
                    'Next i
                End If



                If (DataCombo.Name = "dbcJSubLinea") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codSubLinea"
                    DataCombo.DisplayMember = "descSubLinea"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codSubLinea").ToString()
                    'DataCombo.Text = dt.Rows(i)("descSubLinea").ToString()
                    'End With
                    'Next i
                End If



                If (DataCombo.Name = "dbcRModelo") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codModelo"
                    DataCombo.DisplayMember = "descModelo"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codModelo").ToString()
                    'DataCombo.Text = dt.Rows(i)("descModelo").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcVFamilia") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codFamilia"
                    DataCombo.DisplayMember = "descFamilia"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codFamilia").ToString()
                    'DataCombo.Text = dt.Rows(i)("descFamilia").ToString()
                    'End With
                    'Next i
                End If

                If (DataCombo.Name = "dbcVLinea") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codLinea"
                    DataCombo.DisplayMember = "descLinea"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codLinea").ToString()
                    'DataCombo.Text = dt.Rows(i)("descLinea").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcRMarca") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codMarca"
                    DataCombo.DisplayMember = "descMarca"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codMarca").ToString()
                    'DataCombo.Text = dt.Rows(i)("descMarca").ToString()
                    'End With
                    'Next i
                End If

                If (DataCombo.Name = "dbcMaterial") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codTipoMaterial"
                    DataCombo.DisplayMember = "descTipoMaterial"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codTipoMaterial").ToString()
                    'DataCombo.Text = dt.Rows(i)("descTipoMaterial").ToString()
                    'End With
                    'Next i
                End If

                If (DataCombo.Name = "dbcVendedor") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codVendedor"
                    DataCombo.DisplayMember = "descVendedor"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codVendedor").ToString()
                    'DataCombo.Text = dt.Rows(i)("descVendedor").ToString()
                    'End With
                    'Next i
                End If

                If (DataCombo.Name = "dbcCliente") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "CodCliente"
                    DataCombo.DisplayMember = "DescCliente"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("CodCliente").ToString()
                    'DataCombo.Text = dt.Rows(i)("DescCliente").ToString()
                    'End With
                    ''Next i
                End If

                If (DataCombo.Name = "dbcTaller") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codTaller"
                    DataCombo.DisplayMember = "descTaller"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codTaller").ToString()
                    'DataCombo.Text = dt.Rows(i)("descTaller").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcCaja") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "CodCaja"
                    DataCombo.DisplayMember = "NumCaja"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("CodCaja").ToString()
                    'DataCombo.Text = dt.Rows(i)("NumCaja").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcTipoReparacioN") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "CodGrupo"
                    DataCombo.DisplayMember = "DescGrupo"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("CodGrupo").ToString()
                    'DataCombo.Text = dt.Rows(i)("DescGrupo").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcFamilia") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codFamilia"
                    DataCombo.DisplayMember = "descFamilia"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codFamilia").ToString()
                    'DataCombo.Text = dt.Rows(i)("descFamilia").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "_dbcFamilia_0" Or DataCombo.Name = "_dbcFamilia_1") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codFamilia"
                    DataCombo.DisplayMember = "descFamilia"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codFamilia").ToString()
                    'DataCombo.Text = dt.Rows(i)("descFamilia").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "_dbcLinea_0" Or DataCombo.Name = "_dbcLinea_1") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codLinea"
                    DataCombo.DisplayMember = "descLinea"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codLinea").ToString()
                    'DataCombo.Text = dt.Rows(i)("descLinea").ToString()
                    'End With
                    'Next i
                End If

                If (DataCombo.Name = "dbcSubLinea") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codSubLinea"
                    DataCombo.DisplayMember = "descSubLinea"
                    'DataCombo.SelectedIndex = 1
                    nCodigo = dt.Rows(i)("codSubLinea").ToString()
                    'DataCombo.Text = dt.Rows(i)("descSubLinea").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcKilates") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codKilates"
                    DataCombo.DisplayMember = "descKilates"
                    'Combo.SelectedIndex = 1 
                    nCodigo = dt.Rows(i)("codKilates").ToString()
                    'DataCombo.Text = dt.Rows(i)("descKilates").ToString()
                    'End With
                    'Next i
                End If

                If (DataCombo.Name = "_dbcMaterial_0" Or DataCombo.Name = "_dbcMaterial_1" Or DataCombo.Name = "_dbcMaterial_2") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codTipoMaterial"
                    DataCombo.DisplayMember = "descTipoMaterial"
                    'Combo.SelectedIndex = 1 
                    nCodigo = dt.Rows(i)("codTipoMaterial").ToString()
                    'DataCombo.Text = dt.Rows(i)("descTipoMaterial").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "_cboUnidad_0" Or DataCombo.Name = "_cboUnidad_1" Or DataCombo.Name = "_cboUnidad_2") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codUnidad"
                    DataCombo.DisplayMember = "descUnidad"
                    'DataCombo.SelectedIndex = 1  
                    nCodigo = dt.Rows(i)("codUnidad").ToString()
                    'DataCombo.Text = dt.Rows(i)("descUnidad").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "_cboAlmacen_0" Or DataCombo.Name = "_cboAlmacen_1" Or DataCombo.Name = "_cboAlmacen_2" Or DataCombo.Name = "dbcOrigen" Or DataCombo.Name = "dbcOrigen1") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codAlmacenOrigen"
                    DataCombo.DisplayMember = "descAlmacenOrigen"
                    'DataCombo.SelectedIndex = 1  
                    nCodigo = dt.Rows(i)("codAlmacenOrigen").ToString()
                    'DataCombo.Text = dt.Rows(i)("descAlmacenOrigen").ToString()
                    'End With
                    'Next i
                End If



                If (DataCombo.Name = "_dbcProveedor_0" Or DataCombo.Name = "_dbcProveedor_1" Or DataCombo.Name = "_dbcProveedor_2") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codProvAcreed"
                    DataCombo.DisplayMember = "descProvAcreed"
                    'DataCombo.SelectedIndex = 1  
                    nCodigo = dt.Rows(i)("codProvAcreed").ToString()
                    'DataCombo.Text = dt.Rows(i)("descProvAcreed").ToString()
                    'End With
                    'Next i
                End If



                If (DataCombo.Name = "dbcOrigen") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "CodAlmacenOrigen"
                    DataCombo.DisplayMember = "DescAlmacen"
                    'DataCombo.SelectedIndex = 1  
                    nCodigo = dt.Rows(i)("CodAlmacenOrigen").ToString()
                    'DataCombo.Text = dt.Rows(i)("DescAlmacen").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcMarca") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codMarca"
                    DataCombo.DisplayMember = "descMarca"
                    'DataCombo.SelectedIndex = 1   
                    nCodigo = dt.Rows(i)("codMarca").ToString()
                    'DataCombo.Text = dt.Rows(i)("descMarca").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcModelo") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codModelo"
                    DataCombo.DisplayMember = "descModelo"
                    'DataCombo.SelectedIndex = 1   
                    nCodigo = dt.Rows(i)("codModelo").ToString()
                    'DataCombo.Text = dt.Rows(i)("descModelo").ToString()
                    'End With
                    'Next i
                End If



                If (DataCombo.Name = "_dbcGrupos_0" Or DataCombo.Name = "_dbcGrupos_1") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codUsuario"
                    DataCombo.DisplayMember = "Nombre"
                    'DataCombo.SelectedIndex = 1   
                    nCodigo = dt.Rows(i)("codUsuario").ToString()
                    ' DataCombo.Text = dt.Rows(i)("Nombre").ToString()
                    'End With
                    'Next i
                End If



                If (DataCombo.Name = "dbcModulo") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codModulo"
                    DataCombo.DisplayMember = "descModulo"
                    'DataCombo.SelectedIndex = 1   
                    nCodigo = dt.Rows(i)("codModulo").ToString()
                    ' DataCombo.Text = dt.Rows(i)("descModulo").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcUsuarios") Then
                    'Cmd.CommandTimeout = 1200
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    'For i = 0 To dt.Rows.Count - 1
                    'With DataCombo
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "codUsuario"
                    DataCombo.DisplayMember = "Nombre"
                    'DataCombo.SelectedIndex = 1   
                    nCodigo = dt.Rows(i)("codUsuario").ToString()
                    ' DataCombo.Text = dt.Rows(i)("Nombre").ToString()
                    'End With
                    'Next i
                End If


                If (DataCombo.Name = "dbcAgrupador") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "CodOrigenAplicR"
                    DataCombo.DisplayMember = "DescOrigenAplicR"
                    nCodigo = dt.Rows(i)("CodOrigenAplicR").ToString()
                End If

                If (DataCombo.Name = "dbcRubro") Then
                    Dim dt As New DataTable()
                    Dim adapter As SqlDataAdapter = New SqlDataAdapter(Sql, conexionLocalCliente)
                    adapter.Fill(dt)
                    DataCombo.DataSource = dt
                    DataCombo.ValueMember = "CodRubro"
                    DataCombo.DisplayMember = "DescRubro"
                    nCodigo = dt.Rows(i)("CodRubro").ToString()
                End If



            End If
            'Errores:
        Catch ex As Exception
            If Err.Number <> 0 Then ModErrores.Errores()
        End Try
    End Sub

End Module