Imports XL = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Runtime.InteropServices

Module Module1
    Public xlAPP As New XL.Application
    Public Nombre As String = Nothing
    Sub Main()
        Dim MadrugandoAlAMLO As String = LinkDeHoy()
        Dim Dktp As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim Aqui As String = Dktp & "\Los otros datos"
        Dim Aca As String = Aqui & "\Segun AMLO"
        Dim WkBk As String = Nothing
        Dim Rpte As String = Nothing
        Dim r1 As Integer = Nothing
        Dim c1 As Integer = Nothing
        Dim r2 As Integer = Nothing
        Dim c2 As Integer = Nothing
        Dim j As Integer = Nothing
        Dim chrt As String = Nothing
        Dim muestra As Integer = 0
        Dim fechaSTR As String = Nothing
        Dim eq As String = Nothing

        'Crea directorio de archivos CSV
        If System.IO.Directory.Exists(Aqui) = False Then
            System.IO.Directory.CreateDirectory(Aqui)
            System.IO.Directory.CreateDirectory(Aca)
        Else : End If
        LosOtrosDatos(MadrugandoAlAMLO, Aca & "\" & Nombre & ".csv")
        xlAPP.DisplayAlerts = False
        xlAPP.Workbooks.Open(Aca & "\" & Nombre & ".csv")
        WkBk = xlAPP.ActiveWorkbook.Name

        If System.IO.Directory.Exists(Aqui & "\Resumen de la mañanera.xlsx") = False Then
            xlAPP.Workbooks.Add()
            Rpte = xlAPP.ActiveWorkbook.Name
            xlAPP.Workbooks(Rpte).SaveAs(Aqui & "\Resumen de la mañanera", XL.XlFileFormat.xlWorkbookDefault)
            Rpte = xlAPP.ActiveWorkbook.Name
        Else
            Rpte = "Resumen de la mañanera.xlsx"
            xlAPP.Workbooks.Open(Aqui & "\Resumen de la mañanera.xlsx")
        End If

        SplashScreen1.Show()
        r1 = 34
        c1 = 4
        r2 = 2
        c2 = 2

        While IsNumeric(xlAPP.Workbooks(WkBk).Worksheets(1).cells(r1, c1).value) = True
            'Pega muestras
            xlAPP.Workbooks(Rpte).Worksheets(1).cells(r2, c2).value = xlAPP.Workbooks(WkBk).Worksheets(1).cells(r1, c1).value
            'Pega fechas
            xlAPP.Workbooks(WkBk).Worksheets(1).columns(c1).autofit
            For Each i As Char In xlAPP.Workbooks(WkBk).Worksheets(1).cells(r1 - 33, c1).text
                If i = "-" Then
                    fechaSTR = fechaSTR & "/"
                Else
                    fechaSTR = fechaSTR + i
                End If
            Next
            'fechaSTR = xlAPP.Workbooks(WkBk).Worksheets(1).cells(r1 - 33, c1).text
            xlAPP.Workbooks(Rpte).Worksheets(1).cells(r2, c2 + 1).value = fechaSTR
            'xlAPP.Workbooks(Rpte).Worksheets(1).cells(r2, c2 + 1).NumberFormat = "mm/dd/yyyy"
            'Pega numero de muestras
            muestra = muestra + 1
            xlAPP.Workbooks(Rpte).Worksheets(1).cells(r2, c2 - 1).value = CStr(muestra)
            c1 = c1 + 1
            r2 = r2 + 1
            fechaSTR = Nothing
        End While

        'Agrega encabezados
        With xlAPP.Workbooks(Rpte).Worksheets(1).cells(1, c2 + 1)
            .value = "FECHA"
            .font.bold = True
            .font.color = RGB(206, 17, 38)
        End With
        xlAPP.Workbooks(Rpte).Worksheets(1).columns(c2 + 1).autofit

        With xlAPP.Workbooks(Rpte).Worksheets(1).cells(1, c2)
            .value = "CONFIRMADOS DEL DIA"
            .font.bold = True
            .font.color = RGB(206, 17, 38)
        End With
        xlAPP.Workbooks(Rpte).Worksheets(1).columns(c2).autofit

        With xlAPP.Workbooks(Rpte).Worksheets(1).cells(1, c2 - 1)
            .value = "MUESTRA"
            .font.bold = True
            .font.color = RGB(206, 17, 38)
        End With
        xlAPP.Workbooks(Rpte).Worksheets(1).columns(c2 - 1).autofit
        xlAPP.Workbooks(WkBk).Close()

        'GRAFICA
        With xlAPP.Workbooks(Rpte).Worksheets(1).Shapes.AddChart2(240, XL.XlChartType.xlLine)
            .Name = "Casos confirmados de COVID por dia"
            .chart.setsourcedata(xlAPP.Range("$B$2:$B$" & CStr(muestra + 1)))
            .chart.FullSeriesCollection(1).Trendlines.add(XL.XlTrendlineType.xlPolynomial)
            .chart.Axes(XL.XlAxisType.xlCategory).MajorGridlines.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
            With .chart.FullSeriesCollection(1)
                .format.line.weight = 1
                .format.line.forecolor.rgb = RGB(192, 192, 192)
            End With
            With .chart.FullSeriesCollection(1).Trendlines(1)
                .format.line.forecolor.rgb = RGB(0, 104, 71)
                .format.line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
                .DisplayEquation = True
                eq = .DataLabel.text 'string para guardar el formato de la ecuacion de regresion
            End With
        End With

        xlAPP.Workbooks(Rpte).Save()
        'xlAPP.Visible = True
        'xlAPP.ScreenUpdating = True
        SplashScreen1.Close()
        xlAPP.Quit()
        ReleaseObject(xlAPP)

        MsgBox("Abriendo folder con los datos")
        Process.Start(Aqui)
        'While System.IO.File.Open(Aqui & "\" & Rpte, IO.FileMode.Open) = System.IO.IOException.Equals(0)
        'Dim fi As New FileInfo(Aqui & "\" & Rpte)
        'IsFileOpen(fi)

        'Dim b As Boolean = False

        'While b = False
        '    Try
        '        System.IO.File.Open(Aqui & "\" & Rpte, IO.FileMode.Open, FileAccess.ReadWrite, FileShare.None)
        '    Catch ex As Exception
        '        If TypeOf ex Is IOException Then
        '            'xlAPP.Quit()
        '            ReleaseObject(xlAPP)
        '        Else
        '            b = True
        '        End If
        '    End Try
        'End While
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        'https://stackoverflow.com/questions/15697282/application-not-quitting-after-calling-quit
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Function LinkDeHoy() As String
        Dim fechaCSV As String
        Dim annus As String = Year(Today)
        Dim dies As String = Day(Today) - 1
        Dim mensis As String = Month(Today)

        If Len(mensis) = 1 Then
            mensis = "0" & mensis
        End If

        If Len(dies) = 1 Then
            dies = "0" & dies
        End If

        fechaCSV = annus & mensis & dies
        Nombre = "Casos_Diarios_Estado_Nacional_Confirmados_" & fechaCSV

        LinkDeHoy = "https://coronavirus.gob.mx/datos/Downloads/Files/Casos_Diarios_Estado_Nacional_Confirmados_" & fechaCSV & ".csv"
    End Function

    Sub LosOtrosDatos(linkdelcovixd As String, GuardarAqui As String)
        Dim wc As New Net.WebClient
        wc.DownloadFile(linkdelcovixd, GuardarAqui)
    End Sub

    'Private Sub IsFileOpen(ByVal file As FileInfo)
    '    'https://stackoverflow.com/questions/11287502/vb-net-checking-if-a-file-Is-open-before-proceeding-with-a-read-write
    '    Dim stream As FileStream = Nothing
    '    Dim b As Boolean = False
    '    While b = True
    '        Try
    '            stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
    '            stream.Close()
    '        Catch ex As Exception
    '            If TypeOf ex Is IOException AndAlso IsFileLocked(ex) Then
    '                b = True
    '                MsgBox("Ejte...dehame...pienzo...ehm...io...temgo...otroz...datos")
    '            Else
    '                b = False
    '            End If
    '        End Try
    '    End While
    'End Sub

    'Private Function IsFileLocked(exception As Exception) As Boolean
    '    '   https://stackoverflow.com/questions/11287502/vb-net-checking-if-a-file-is-open-before-proceeding-with-a-read-write
    '    Dim errorCode As Integer = Marshal.GetHRForException(exception) And ((1 << 16) - 1)
    '    Return errorCode = 32 OrElse errorCode = 33
    'End Function
End Module

