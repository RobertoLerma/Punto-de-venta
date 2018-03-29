'**********************************************************************************************************************'
'*PROGRAMA: MODULO DE FRECUENCIA JOYERIA RAMOS  
'*AUTOR: MIGUEL ANGEL GARCIA WHA 
'*EMPRESA: GRUPO VITEK
'*FECHA DE INICIO: 02/01/2018     
'*FECHA DE TERMINACION:   
'**********************************************************************************************************************'

Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Public Module ModFrecuencia

    'UPGRADE_WARNING: Lower bound of array aFechasFrecuencia was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Public aFechasFrecuencia(200000) As Date

    '-----------------------------------------------------------------------------------------------------------------
    ' CONSTANTES PARA EL CÁLCULO DE LA FRECUENCIA MENSUAL Y ANUAL
    '-----------------------------------------------------------------------------------------------------------------
    Const C_PRIMER As Integer = 0
    Const C_SEGUNDO As Integer = 1
    Const C_TERCER As Integer = 2
    Const C_CUARTO As Integer = 3
    Const C_ULTIMO As Integer = 4

    Const C_DIA As Integer = 0
    Const C_DIASEMANA As Integer = 1
    Const C_DIAFINSEMANA As Integer = 2
    Const C_DOMINGO As Integer = 3
    Const C_LUNES As Integer = 4
    Const C_MARTES As Integer = 5
    Const C_MIERCOLES As Integer = 6
    Const C_JUEVES As Integer = 7
    Const C_VIERNES As Integer = 8
    Const C_SABADO As Integer = 9

    'UPGRADE_WARNING: Lower bound of array aMeses was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim aMeses(12) As Integer

    '-----------------------------------------------------------------------------------------------------------------
    ' FINALIZAN CONSTANTES PARA EL CÁLCULO DE LA FRECUENCIA MENSUAL Y ANUAL
    '-----------------------------------------------------------------------------------------------------------------

    Dim cFrecuencia As String
    Dim nTipoIntervalo As Integer
    Public nRepeticiones As Integer
    Dim dFechaInicio As Date
    Dim dFechaFin As Date
    Dim nPeriodo As Integer
    Dim cDiaSemana As String
    Dim nDiaMes As Integer
    Dim nMes As Integer
    Dim nOpcion As Integer
    Dim nCual As Integer
    Dim nCuando As Integer

    Dim cMsg As String

    Public Sub GenerarFrecuencia(ByRef pFrecuencia As String, ByRef pTipoIntervalo As Integer, ByRef pRepeticiones As Integer, ByRef pFechaInicio As Date, ByRef pFechaFin As Date, ByRef pPeriodo As Integer, ByRef pDiaSemana As String, ByRef pDiaMes As Integer, ByRef pMes As Integer, ByRef pOpcion As Integer, ByRef pCual As Integer, ByRef pCuando As Integer)
        Dim nAnio As Integer

        cFrecuencia = pFrecuencia
        nTipoIntervalo = pTipoIntervalo
        nRepeticiones = pRepeticiones
        dFechaInicio = pFechaInicio
        dFechaFin = pFechaFin
        nPeriodo = pPeriodo
        cDiaSemana = pDiaSemana
        nDiaMes = pDiaMes
        nMes = pMes
        nOpcion = pOpcion
        nCual = pCual
        nCuando = pCuando

        nAnio = Year(dFechaInicio)

        Call InicializaVectores(nAnio)

        Select Case pFrecuencia
            Case "0" ' Diaria
                GenerarFrecuenciaDiaria()
            Case "1" ' Semanal
                GenerarFrecuenciaSemanal()
            Case "2" ' Mensual
                GenerarFrecuenciaMensual()
            Case "3" ' Anual
                GenerarFrecuenciaAnual()
        End Select
        'MsgBox cMsg, vbOKOnly + vbInformation, gstrNombCortoEmpresa
    End Sub

    Public Sub GenerarFrecuenciaDiaria()
        Dim nDiasTotales As Integer
        Dim dFechaAux As Date
        Dim I As Integer

        nDiasTotales = dFechaFin.ToOADate - dFechaInicio.ToOADate
        nRepeticiones = 0
        nRepeticiones = CShort(Format(nDiasTotales / nPeriodo, "####0"))
        dFechaAux = dFechaInicio
        cMsg = ""
        For I = 1 To nRepeticiones
            cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
            aFechasFrecuencia(I) = dFechaAux
            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + nPeriodo)
        Next I
    End Sub

    Public Sub GenerarFrecuenciaSemanal()
        Dim nContador As Integer
        Dim nDiaMarcado As Integer
        Dim dFechaAux As Date
        Dim I As Integer
        Dim nSemanas As Integer

        nSemanas = nPeriodo
        nRepeticiones = 0
        cMsg = ""
        nContador = 1
        Do While nContador <= Len(Trim(cDiaSemana))
            nDiaMarcado = CShort(Numerico(Mid(cDiaSemana, nContador, 1)))
            nContador = nContador + 1
            'Encontrar la primer ocurrencia de este día, a partir de la fecha de inicio
            dFechaAux = dFechaInicio
            Do While nDiaMarcado <> Weekday(dFechaAux)
                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
            Loop
            nPeriodo = nSemanas * 7
            Do While dFechaAux <= dFechaFin
                cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                nRepeticiones = nRepeticiones + 1
                aFechasFrecuencia(nRepeticiones) = dFechaAux
                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + nPeriodo)
            Loop
        Loop
    End Sub

    Public Sub GenerarFrecuenciaMensual()
        Dim nAnio As Integer
        Dim nDia As Integer
        Dim nDiaSemana As Integer
        Dim nOcurrencia As Integer
        Dim nContador As Integer
        Dim I As Integer

        Dim dFechaAux As Date
        Dim cFechaAux As String

        nRepeticiones = 0
        Select Case nOpcion
            Case 0
                dFechaAux = dFechaInicio
                nMes = Month(dFechaInicio)
                nAnio = Year(dFechaInicio)
                cMsg = ""
                Do While dFechaAux <= dFechaFin
                    Call InicializaVectores(nAnio)
                    If nDiaMes > aMeses(nMes) Then
                        nDia = aMeses(nMes)
                        cFechaAux = CStr(nDia) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                        dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                    Else
                        cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                        dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                        If dFechaAux <= dFechaFin And dFechaAux >= dFechaInicio Then
                            cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                            nRepeticiones = nRepeticiones + 1
                            aFechasFrecuencia(nRepeticiones) = dFechaAux
                        End If
                    End If
                    For I = 1 To nPeriodo
                        nMes = nMes + 1
                        If nMes > 12 Then
                            nMes = 1
                            nAnio = nAnio + 1
                        End If
                    Next I
                Loop
            Case 1
                'Determina el día del mes
                Select Case nCuando
                    Case C_DIA
                        If nCual < C_ULTIMO Then
                            nDiaMes = nCual + 1
                            'Encontrar la primer ocurrencia de este día, a partir de la fecha de inicio
                            dFechaAux = dFechaInicio
                            Do While nDiaMes <> VB.Day(dFechaAux)
                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                            Loop
                            dFechaInicio = dFechaAux

                            nMes = Month(dFechaInicio)
                            nAnio = Year(dFechaInicio)
                            cMsg = ""
                            nRepeticiones = 0
                            Do While dFechaAux <= dFechaFin
                                Call InicializaVectores(nAnio)
                                If nDiaMes > aMeses(nMes) Then
                                    nDia = aMeses(nMes)
                                    cFechaAux = CStr(nDia) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Else
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    If dFechaAux <= dFechaFin Then
                                        cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                        nRepeticiones = nRepeticiones + 1
                                        aFechasFrecuencia(nRepeticiones) = dFechaAux
                                    End If
                                End If
                                For I = 1 To nPeriodo
                                    nMes = nMes + 1
                                    If nMes > 12 Then
                                        nMes = 1
                                        nAnio = nAnio + 1
                                    End If
                                Next I
                            Loop
                        Else 'Es el último día del mes
                            nDiaMes = aMeses(Month(dFechaInicio))
                            'Encontrar la primer ocurrencia de este día, a partir de la fecha de inicio
                            dFechaAux = dFechaInicio
                            Do While nDiaMes <> VB.Day(dFechaAux)
                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                            Loop
                            dFechaInicio = dFechaAux
                            nMes = Month(dFechaInicio)
                            nAnio = Year(dFechaInicio)
                            cMsg = ""
                            nRepeticiones = 0
                            Do While dFechaAux <= dFechaFin
                                Call InicializaVectores(nAnio)
                                nDiaMes = aMeses(nMes)

                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                If dFechaAux <= dFechaFin Then
                                    cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                    nRepeticiones = nRepeticiones + 1
                                    aFechasFrecuencia(nRepeticiones) = dFechaAux
                                End If

                                For I = 1 To nPeriodo
                                    nMes = nMes + 1
                                    If nMes > 12 Then
                                        nMes = 1
                                        nAnio = nAnio + 1
                                    End If
                                Next I
                            Loop
                        End If
                    Case C_DIASEMANA
                        If nCual < C_ULTIMO Then
                            nDiaSemana = nCual + 1
                        Else 'Es el último día de la primer semana del mes
                            nDiaSemana = FirstDayOfWeek.Saturday
                        End If
                        'Encontrar la primera ocurrencia
                        'Después del día 7 del mes, ya sucedieron todos los días de la semana
                        dFechaAux = dFechaInicio
                        nMes = Month(dFechaAux)
                        nAnio = Year(dFechaAux)
                        If VB.Day(dFechaAux) > 7 Then
                            nDiaMes = 1
                            nMes = nMes + 1
                            If nMes > 12 Then
                                nMes = 1
                                nAnio = nAnio + 1
                            End If
                            cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                            dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                            Do While Weekday(dFechaAux) <> nDiaSemana
                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                            Loop
                        Else
                            Do While Weekday(dFechaAux) <> nDiaSemana
                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                            Loop
                            If VB.Day(dFechaAux) > 7 Then
                                nDiaMes = 1
                                nMes = nMes + 1
                                If nMes > 12 Then
                                    nMes = 1
                                    nAnio = nAnio + 1
                                End If
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Do While Weekday(dFechaAux) <> nDiaSemana
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                Loop
                            End If
                        End If
                        'Aquí ya tengo la primer ocurrencia
                        dFechaInicio = dFechaAux
                        cMsg = ""
                        nRepeticiones = 0
                        Do While dFechaAux <= dFechaFin
                            cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                            nRepeticiones = nRepeticiones + 1
                            aFechasFrecuencia(nRepeticiones) = dFechaAux
                            nDiaMes = 1
                            For I = 1 To nPeriodo
                                nMes = nMes + 1
                                If nMes > 12 Then
                                    nMes = 1
                                    nAnio = nAnio + 1
                                End If
                            Next I
                            cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                            dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                            Do While Weekday(dFechaAux) <> nDiaSemana
                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                            Loop
                        Loop
                    Case C_DIAFINSEMANA
                        nDiaSemana = FirstDayOfWeek.Saturday
                        If nCual < C_ULTIMO Then
                            nOcurrencia = nCual + 1
                            'Obtener la primera ocurrencia
                            dFechaAux = dFechaInicio
                            nMes = Month(dFechaAux)
                            nAnio = Year(dFechaAux)
                            nDiaMes = 1
                            cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                            dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                            nContador = 0
                            Do While nContador < nOcurrencia
                                If Weekday(dFechaAux) <> FirstDayOfWeek.Saturday And Weekday(dFechaAux) <> FirstDayOfWeek.Sunday Then
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                ElseIf Weekday(dFechaAux) = FirstDayOfWeek.Saturday Or Weekday(dFechaAux) = FirstDayOfWeek.Sunday Then
                                    nContador = nContador + 1
                                    If nContador < nOcurrencia Then
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                    End If
                                End If
                            Loop
                            If dFechaAux < dFechaInicio Then
                                'Pasa a buscar al siguiente mes
                                nDiaMes = 1
                                nMes = nMes + 1
                                If nMes > 12 Then
                                    nMes = 1
                                    nAnio = nAnio + 1
                                End If
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                nContador = 0
                                Do While nContador < nOcurrencia
                                    If Weekday(dFechaAux) <> FirstDayOfWeek.Saturday And Weekday(dFechaAux) <> FirstDayOfWeek.Sunday Then
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                    ElseIf Weekday(dFechaAux) = FirstDayOfWeek.Saturday Or Weekday(dFechaAux) = FirstDayOfWeek.Sunday Then
                                        nContador = nContador + 1
                                        If nContador < nOcurrencia Then
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    End If
                                Loop
                            End If
                            'Aquí ya se tiene la primer ocurrencia
                            dFechaInicio = dFechaAux
                            cMsg = ""
                            nRepeticiones = 0
                            Do While dFechaAux <= dFechaFin
                                cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                nRepeticiones = nRepeticiones + 1
                                aFechasFrecuencia(nRepeticiones) = dFechaAux
                                nDiaMes = 1
                                For I = 1 To nPeriodo
                                    nMes = nMes + 1
                                    If nMes > 12 Then
                                        nMes = 1
                                        nAnio = nAnio + 1
                                    End If
                                Next I
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                nContador = 0
                                Do While nContador < nOcurrencia
                                    If Weekday(dFechaAux) <> FirstDayOfWeek.Saturday And Weekday(dFechaAux) <> FirstDayOfWeek.Sunday Then
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                    ElseIf Weekday(dFechaAux) = FirstDayOfWeek.Saturday Or Weekday(dFechaAux) = FirstDayOfWeek.Sunday Then
                                        nContador = nContador + 1
                                        If nContador < nOcurrencia Then
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    End If
                                Loop
                            Loop
                        Else 'Es el C_ULTIMO día del fin de semana (Sábado o Domingo)
                            'Obtener la primera ocurrencia
                            nDiaSemana = FirstDayOfWeek.Saturday
                            nAnio = Year(dFechaInicio)
                            nMes = Month(dFechaInicio)
                            Call InicializaVectores(nAnio)
                            dFechaAux = dFechaInicio
                            nDiaMes = aMeses(nMes)
                            cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                            dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                            Do While Weekday(dFechaAux) <> FirstDayOfWeek.Saturday And Weekday(dFechaAux) <> FirstDayOfWeek.Sunday
                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                            Loop
                            If dFechaAux < dFechaInicio Then
                                'Pasar al siguiente mes
                                nMes = nMes + 1
                                If nMes > 12 Then
                                    nMes = 1
                                    nAnio = nAnio + 1
                                End If
                                Call InicializaVectores(nAnio)
                                nDiaMes = aMeses(nMes)
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Do While Weekday(dFechaAux) <> FirstDayOfWeek.Saturday And Weekday(dFechaAux) <> FirstDayOfWeek.Sunday
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                Loop
                            End If
                            'Aquí ya se tiene la primer ocurrencia
                            dFechaInicio = dFechaAux
                            cMsg = ""
                            nRepeticiones = 0
                            Do While dFechaAux <= dFechaFin
                                cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                nRepeticiones = nRepeticiones + 1
                                aFechasFrecuencia(nRepeticiones) = dFechaAux
                                For I = 1 To nPeriodo
                                    nMes = nMes + 1
                                    If nMes > 12 Then
                                        nMes = 1
                                        nAnio = nAnio + 1
                                    End If
                                Next I
                                Call InicializaVectores(nAnio)
                                nDiaMes = aMeses(nMes)
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Do While Weekday(dFechaAux) <> FirstDayOfWeek.Saturday And Weekday(dFechaAux) <> FirstDayOfWeek.Sunday
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                Loop
                            Loop
                        End If
                    Case Else 'Está entre C_DOMINGO Y C_SABADO
                        nDiaSemana = nCuando - 2
                        If nCual < C_ULTIMO Then
                            nOcurrencia = nCual + 1
                            'Obtener la primera ocurrencia
                            dFechaAux = dFechaInicio
                            nMes = Month(dFechaAux)
                            nAnio = Year(dFechaAux)
                            nDiaMes = 1
                            cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                            dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                            nContador = 0
                            Do While nContador < nOcurrencia
                                If Weekday(dFechaAux) <> nDiaSemana Then
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                ElseIf Weekday(dFechaAux) = nDiaSemana Then
                                    nContador = nContador + 1
                                    If nContador < nOcurrencia Then
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                    End If
                                End If
                            Loop
                            If dFechaAux < dFechaInicio Then
                                'Pasa a buscar al siguiente mes
                                nDiaMes = 1
                                nMes = nMes + 1
                                If nMes > 12 Then
                                    nMes = 1
                                    nAnio = nAnio + 1
                                End If
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                nContador = 0
                                Do While nContador < nOcurrencia
                                    If Weekday(dFechaAux) <> nDiaSemana Then
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                    ElseIf Weekday(dFechaAux) = nDiaSemana Then
                                        nContador = nContador + 1
                                        If nContador < nOcurrencia Then
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    End If
                                Loop
                            End If
                            'Aquí ya se tiene la primer ocurrencia
                            dFechaInicio = dFechaAux
                            cMsg = ""
                            nRepeticiones = 0
                            Do While dFechaAux <= dFechaFin
                                cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                nRepeticiones = nRepeticiones + 1
                                aFechasFrecuencia(nRepeticiones) = dFechaAux
                                nDiaMes = 1
                                For I = 1 To nPeriodo
                                    nMes = nMes + 1
                                    If nMes > 12 Then
                                        nMes = 1
                                        nAnio = nAnio + 1
                                    End If
                                Next I
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                nContador = 0
                                Do While nContador < nOcurrencia
                                    If Weekday(dFechaAux) <> nDiaSemana Then
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                    ElseIf Weekday(dFechaAux) = nDiaSemana Then
                                        nContador = nContador + 1
                                        If nContador < nOcurrencia Then
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    End If
                                Loop
                            Loop
                        Else 'Es el C_ULTIMO día especificado, en el mes
                            'Obtener la primera ocurrencia
                            nDiaSemana = nCuando - 2
                            nAnio = Year(dFechaInicio)
                            nMes = Month(dFechaInicio)
                            Call InicializaVectores(nAnio)
                            dFechaAux = dFechaInicio
                            nDiaMes = aMeses(nMes)
                            cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                            dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                            Do While Weekday(dFechaAux) <> nDiaSemana
                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                            Loop
                            If dFechaAux < dFechaInicio Then
                                'Pasar al siguiente mes
                                nMes = nMes + 1
                                If nMes > 12 Then
                                    nMes = 1
                                    nAnio = nAnio + 1
                                End If
                                Call InicializaVectores(nAnio)
                                nDiaMes = aMeses(nMes)
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Do While Weekday(dFechaAux) <> nDiaSemana
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                Loop
                            End If
                            'Aquí ya se tiene la primer ocurrencia
                            dFechaInicio = dFechaAux
                            cMsg = ""
                            nRepeticiones = 0
                            Do While dFechaAux <= dFechaFin
                                cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                nRepeticiones = nRepeticiones + 1
                                aFechasFrecuencia(nRepeticiones) = dFechaAux
                                For I = 1 To nPeriodo
                                    nMes = nMes + 1
                                    If nMes > 12 Then
                                        nMes = 1
                                        nAnio = nAnio + 1
                                    End If
                                Next I
                                Call InicializaVectores(nAnio)
                                nDiaMes = aMeses(nMes)
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMes) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Do While Weekday(dFechaAux) <> nDiaSemana
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                Loop
                            Loop
                        End If
                End Select
        End Select
    End Sub

    Public Sub GenerarFrecuenciaAnual()
        Dim nDiaSemana As Integer
        Dim nDia As Integer
        Dim nAnio As Integer
        Dim nMesSeleccionado As Integer
        Dim nContador As Integer
        Dim nOcurrencia As Integer

        Dim dFechaAux As Date
        Dim cFechaAux As String

        nMesSeleccionado = nMes

        Select Case nOpcion
            Case 0
                'Obtener la primera ocurrencia
                dFechaAux = dFechaInicio

                nMes = Month(dFechaAux)
                nDia = VB.Day(dFechaAux)
                nAnio = Year(dFechaAux)

                If nMesSeleccionado <> nMes Then
                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                    If dFechaAux < dFechaInicio Then
                        nAnio = nAnio + 1
                        cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                    Else 'If dFechaAux >= dFechaIni Then
                        cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                    End If
                Else 'nMesSeleccionado = nMes
                    If nDia <= nDiaMes Then
                        cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                    Else
                        nAnio = nAnio + 1
                        cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                    End If
                End If
                'Aquí ya se tiene la fecha de la primer ocurrencia
                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                dFechaInicio = dFechaAux
                cMsg = ""
                nRepeticiones = 0
                Do While dFechaAux <= dFechaFin
                    cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                    nRepeticiones = nRepeticiones + 1
                    aFechasFrecuencia(nRepeticiones) = dFechaAux
                    nAnio = nAnio + 1
                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                Loop
            Case 1
                Select Case nCuando
                    Case C_DIA
                        Select Case True
                            Case nCual < C_ULTIMO
                                nDiaMes = nCual + 1
                                'Encontrar la primer ocurrencia de este día, a partir de la fecha de inicio
                                dFechaAux = dFechaInicio

                                nMes = Month(dFechaAux)
                                nDia = VB.Day(dFechaAux)
                                nAnio = Year(dFechaAux)

                                If nMesSeleccionado <> nMes Then
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    If dFechaAux < dFechaInicio Then
                                        nAnio = nAnio + 1
                                        cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    Else 'If dFechaAux >= dFechaIni Then
                                        cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    End If
                                Else 'nMesSeleccionado = nMes
                                    If nDia <= nDiaMes Then
                                        cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    Else
                                        nAnio = nAnio + 1
                                        cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    End If
                                End If
                                'Aquí ya tengo la primer ocurrencia
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                dFechaInicio = dFechaAux
                                cMsg = ""
                                nRepeticiones = 0
                                Do While dFechaAux <= dFechaFin
                                    cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                    nRepeticiones = nRepeticiones + 1
                                    aFechasFrecuencia(nRepeticiones) = dFechaAux
                                    nAnio = nAnio + 1
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Loop
                            Case nCual = C_ULTIMO
                                nDiaMes = aMeses(nMesSeleccionado)
                                nAnio = Year(dFechaInicio)
                                'Encontrar la primer ocurrencia de este día, a partir de la fecha de inicio
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                cMsg = ""
                                nRepeticiones = 0
                                Do While dFechaAux <= dFechaFin
                                    cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                    nRepeticiones = nRepeticiones + 1
                                    aFechasFrecuencia(nRepeticiones) = dFechaAux
                                    nAnio = nAnio + 1
                                    Call InicializaVectores(nAnio)
                                    nDiaMes = aMeses(nMesSeleccionado)
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Loop
                        End Select
                    Case C_DIASEMANA
                        Select Case True
                            Case nCual < C_ULTIMO
                                nDiaSemana = nCual + 1
                            Case Else 'C_ULTIMO día de la primer semana del mes
                                nDiaSemana = FirstDayOfWeek.Saturday
                        End Select
                        'Encontrar la primera ocurrencia
                        'Después del día 7 del mes, ya sucedieron todos los días de la semana
                        dFechaAux = dFechaInicio
                        nMes = Month(dFechaAux)
                        nAnio = Year(dFechaAux)
                        nDiaMes = VB.Day(dFechaAux)

                        If nMesSeleccionado <> nMes Then
                            cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                            dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                            If dFechaAux < dFechaInicio Then
                                nAnio = nAnio + 1
                                nDiaMes = 1
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                            Else 'If dFechaAux > dFechaIni Then
                                nDiaMes = 1
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                            End If
                        Else 'nMesSeleccionado = nmes
                            cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                            dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                            If VB.Day(dFechaAux) > 7 Then
                                nAnio = nAnio + 1
                                nDiaMes = 1
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                            Else
                                'Busca el día seleccionado desde el día actual
                                Do While Weekday(dFechaAux) <> nDiaSemana
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                Loop
                                'Si al encontrarlo la fecha es mayor que la de inicio, incrementa el año
                                If dFechaAux > dFechaInicio Then
                                    nAnio = nAnio + 1
                                    nDiaMes = 1
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                Else
                                    nDiaMes = 1
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                End If
                            End If
                        End If
                        dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                        Do While Weekday(dFechaAux) <> nDiaSemana
                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                        Loop
                        'Aquí ya tengo la primer ocurrencia
                        dFechaInicio = dFechaAux
                        cMsg = ""
                        nRepeticiones = 0
                        Do While dFechaAux <= dFechaFin
                            cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                            nRepeticiones = nRepeticiones + 1
                            aDiasPago(nRepeticiones) = dFechaAux
                            nDiaMes = 1
                            nAnio = nAnio + 1
                            cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                            dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                            Do While Weekday(dFechaAux) <> nDiaSemana
                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                            Loop
                        Loop
                    Case C_DIAFINSEMANA
                        Select Case True
                            Case nCual < C_ULTIMO
                                nOcurrencia = nCual + 1
                                dFechaAux = dFechaInicio
                                nMes = Month(dFechaAux)
                                nAnio = Year(dFechaAux)

                                'Obtener la primera ocurrencia en el mes actual
                                nDiaMes = 1
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))

                                nContador = 0
                                Do While nContador < nOcurrencia
                                    If Weekday(dFechaAux) = FirstDayOfWeek.Saturday Or Weekday(dFechaAux) = FirstDayOfWeek.Sunday Then
                                        nContador = nContador + 1
                                        If nContador < nOcurrencia Then
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    Else
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                    End If
                                Loop
                                If dFechaAux < dFechaInicio Then
                                    nDiaMes = 1
                                    nAnio = nAnio + 1
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    nContador = 0
                                    Do While nContador < nOcurrencia
                                        If Weekday(dFechaAux) = FirstDayOfWeek.Saturday Or Weekday(dFechaAux) = FirstDayOfWeek.Sunday Then
                                            nContador = nContador + 1
                                            If nContador < nOcurrencia Then
                                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                            End If
                                        Else
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    Loop
                                End If
                                'Aquí ya tengo la primer ocurrencia
                                dFechaInicio = dFechaAux
                                cMsg = ""
                                nRepeticiones = 0
                                Do While dFechaAux <= dFechaFin
                                    cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                    nRepeticiones = nRepeticiones + 1
                                    aFechasFrecuencia(nRepeticiones) = dFechaAux
                                    nDiaMes = 1
                                    nAnio = nAnio + 1
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    nContador = 0
                                    Do While nContador < nOcurrencia
                                        If Weekday(dFechaAux) = FirstDayOfWeek.Saturday Or Weekday(dFechaAux) = FirstDayOfWeek.Sunday Then
                                            nContador = nContador + 1
                                            If nContador < nOcurrencia Then
                                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                            End If
                                        Else
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    Loop
                                Loop
                            Case Else 'C_ULTIMO
                                dFechaAux = dFechaInicio
                                nMes = Month(dFechaAux)
                                nAnio = Year(dFechaAux)
                                Call InicializaVectores(nAnio)
                                nDiaMes = aMeses(nMesSeleccionado)
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Do While Weekday(dFechaAux) <> FirstDayOfWeek.Saturday And Weekday(dFechaAux) <> FirstDayOfWeek.Sunday
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                Loop
                                If dFechaAux < dFechaInicio Then
                                    nAnio = nAnio + 1
                                    Call InicializaVectores(nAnio)
                                    nDiaMes = aMeses(nMesSeleccionado)
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    Do While Weekday(dFechaAux) <> FirstDayOfWeek.Saturday And Weekday(dFechaAux) <> FirstDayOfWeek.Sunday
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                    Loop
                                End If
                                'Aquí ya tengo la primer ocurrencia
                                dFechaInicio = dFechaAux
                                cMsg = ""
                                nRepeticiones = 0
                                Do While dFechaAux <= dFechaFin
                                    cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                    nRepeticiones = nRepeticiones + 1
                                    aDiasPago(nRepeticiones) = dFechaAux
                                    nAnio = nAnio + 1
                                    Call InicializaVectores(nAnio)
                                    nDiaMes = aMeses(nMesSeleccionado)
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    Do While Weekday(dFechaAux) <> FirstDayOfWeek.Saturday And Weekday(dFechaAux) <> FirstDayOfWeek.Sunday
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                    Loop
                                Loop
                        End Select
                    Case Else 'Se encuentra entre C_DOMINGO y C_SABADO
                        nDiaSemana = nCuando - 2
                        Select Case True
                            Case nCual < C_ULTIMO
                                nOcurrencia = nCual + 1

                                dFechaAux = dFechaInicio
                                nMes = Month(dFechaAux)
                                nAnio = Year(dFechaAux)

                                'Obtener la primera ocurrencia en el mes actual
                                nDiaMes = 1
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))

                                nContador = 0
                                Do While nContador < nOcurrencia
                                    If Weekday(dFechaAux) = nDiaSemana Then
                                        nContador = nContador + 1
                                        If nContador < nOcurrencia Then
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    Else
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                    End If
                                Loop
                                If dFechaAux < dFechaInicio Then
                                    nDiaMes = 1
                                    nAnio = nAnio + 1
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    nContador = 0
                                    Do While nContador < nOcurrencia
                                        If Weekday(dFechaAux) = nDiaSemana Then
                                            nContador = nContador + 1
                                            If nContador < nOcurrencia Then
                                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                            End If
                                        Else
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    Loop
                                End If
                                'Aquí ya tengo la primera ocurrencia
                                dFechaInicio = dFechaAux
                                cMsg = ""
                                nRepeticiones = 0
                                Do While dFechaAux <= dFechaFin
                                    cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                    nRepeticiones = nRepeticiones + 1
                                    aFechasFrecuencia(nRepeticiones) = dFechaAux
                                    nDiaMes = 1
                                    nAnio = nAnio + 1
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    nContador = 0
                                    Do While nContador < nOcurrencia
                                        If Weekday(dFechaAux) = nDiaSemana Then
                                            nContador = nContador + 1
                                            If nContador < nOcurrencia Then
                                                dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                            End If
                                        Else
                                            dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate + 1)
                                        End If
                                    Loop
                                Loop
                            Case Else 'C_ULTIMO
                                dFechaAux = dFechaInicio
                                nMes = Month(dFechaAux)
                                nAnio = Year(dFechaAux)
                                Call InicializaVectores(nAnio)
                                nDiaMes = aMeses(nMesSeleccionado)
                                cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                Do While Weekday(dFechaAux) <> nDiaSemana
                                    dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                Loop
                                If dFechaAux < dFechaInicio Then
                                    nAnio = nAnio + 1
                                    Call InicializaVectores(nAnio)
                                    nDiaMes = aMeses(nMesSeleccionado)
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    Do While Weekday(dFechaAux) <> nDiaSemana
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                    Loop
                                End If
                                'Aquí ya tengo la primer ocurrencia
                                dFechaInicio = dFechaAux
                                cMsg = ""
                                nRepeticiones = 0
                                Do While dFechaAux <= dFechaFin
                                    cMsg = cMsg & Format(dFechaAux, "dd/MMM/yyyy") & vbNewLine
                                    nRepeticiones = nRepeticiones + 1
                                    aFechasFrecuencia(nRepeticiones) = dFechaAux
                                    nAnio = nAnio + 1
                                    Call InicializaVectores(nAnio)
                                    nDiaMes = aMeses(nMesSeleccionado)
                                    cFechaAux = CStr(nDiaMes) & "/" & CStr(nMesSeleccionado) & "/" & CStr(nAnio)
                                    dFechaAux = CDate(Format(CDate(cFechaAux), "dd/MMM/yyyy"))
                                    Do While Weekday(dFechaAux) <> nDiaSemana
                                        dFechaAux = System.DateTime.FromOADate(dFechaAux.ToOADate - 1)
                                    Loop
                                Loop
                        End Select
                End Select
        End Select
    End Sub

    Public Sub InicializaVectorFecha()
        Dim I As Integer
        For I = 1 To C_DIASPAGO
            aFechasFrecuencia(I) = #1/1/9999#
        Next I
    End Sub

    Public Sub InicializaVectores(ByRef nAnio As Integer)
        If (nAnio Mod 4) = 0 Then
            aMeses(2) = 29 'Febrero
        Else
            aMeses(2) = 28 'Febrero
        End If
        aMeses(1) = 31 'Enero
        aMeses(3) = 31 'Marzo
        aMeses(4) = 30 'Abril
        aMeses(5) = 31 'Mayo
        aMeses(6) = 30 'Junio
        aMeses(7) = 31 'Julio
        aMeses(8) = 31 'Agosto
        aMeses(9) = 30 'Septiembre
        aMeses(10) = 31 'Octubre
        aMeses(11) = 30 'Noviembre
        aMeses(12) = 31 'Diciembre
    End Sub
End Module