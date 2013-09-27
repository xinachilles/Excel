
Imports OfficeOpenXml
Imports System.IO
Imports System.Collections.Generic
Imports System.Drawing

Public Class ExcelSupport
    '''<summary> 
    ''' Export the 2D contour map to excel file 
    '''</summary>
    ''' <param name="data_info">comments.</param>
    '''<param name="path">path the excel file will export to</param>
    '''<param name= "topography">2D contour map</param>
    '''<remarks></remarks>
    Public Sub PastePicToExcel(ByVal data_info As String, ByVal path As String, ByVal topography As Bitmap)
        Try
            If topography Is Nothing = False Then

                Dim work_file As FileInfo = New FileInfo(path)

                Using work_excel As ExcelPackage = New ExcelPackage(work_file)

                    Dim k As Integer = work_excel.Workbook.Worksheets.Count
                    Dim work_sheet As ExcelWorksheet = work_excel.Workbook.Worksheets.Add(k.ToString)

                    Dim pic As OfficeOpenXml.Drawing.ExcelPicture = work_sheet.Drawings.AddPicture(Guid.NewGuid().ToString(), topography)
                    pic.From.Column = 0
                    pic.From.Row = 0
                    ' pic.From.ColumnOff = ExcelHelper.MTU2Pixel(2);
                    ' pic.From.RowOff = ExcelHelper.MTU2Pixel(2);
                    pic.SetSize(topography.Width, topography.Height)
                    work_sheet.Cells(1, 1).Value = data_info
                    work_excel.Save()
                End Using
            End If
        Catch ex As Exception

        End Try
    End Sub
    '''<summary> 
    ''' Export data excel file 
    '''</summary>
    ''' <param name="data_info">comments about data.</param>
    '''<param name="path">path the excel file will export to</param>
    '''<param name= "data">data store in the excel file</param>
    '''<remarks></remarks>
    Public Sub WriteDataToExcel(ByVal data_info As String, ByVal path As String, ByVal data As List(Of Double()))
        Try
            Dim i As Integer = 1
            Dim j As Integer = 0
            'Dim sum As Double
            Dim series As Drawing.Chart.ExcelChartSerie
            Dim d As Double()

            Dim x_series As String
            Dim work_cell As Integer = 1
            ' Dim work_cell_end As Integer
            Dim work_file As FileInfo = New FileInfo(path)

            Using work_excel As ExcelPackage = New ExcelPackage(work_file)


                Dim k As Integer = work_excel.Workbook.Worksheets.Count

                Dim work_sheet As ExcelWorksheet = work_excel.Workbook.Worksheets.Add(k.ToString)
                Dim chart As Drawing.Chart.ExcelChart = work_sheet.Drawings.AddChart(k.ToString, Drawing.Chart.eChartType.XYScatterLinesNoMarkers)
                chart.Legend.Remove()
                ' chart.Title.Text = Nothing




                'chart. 
                While True
                    If work_sheet.Cells.Item(i, 1).Value Is Nothing = True Then
                        work_cell = i
                        Exit While
                    Else
                        i = i + 1
                    End If
                End While
                ' put x_series
                'For i = 1 To data.Item(0).Count
                '    work_sheet.Cells.Item(i, work_cell).Value = i
                'Next
                work_sheet.Cells(1, data.Count + 1).Value = "Min"
                work_sheet.Cells(2, data.Count + 1).Value = "Max"
                work_sheet.Cells(3, data.Count + 1).Value = "Average"
                work_sheet.Cells(4, data.Count + 1).Value = "Std"
                work_sheet.Cells(5, data.Count + 1).Value = data_info





                For Each d In data
                    Dim min As Double = Double.MaxValue
                    Dim max As Double = Double.MinValue

                    If d Is Nothing = False Then
                        j = 1
                        For Each m As Double In d
                            work_sheet.Cells(j, work_cell).Value = m

                            If m < min Then min = m
                            If m > max Then max = m

                            j = j + 1
                        Next

                        If work_cell = 1 Then x_series = ExcelRange.GetAddress(1, 1, data.Item(0).Count, 1)
                        If work_cell >= 2 Then series = chart.Series.Add(ExcelRange.GetAddress(1, work_cell, data.Item(0).Count, work_cell), x_series)

                        'min
                        work_sheet.Cells(1, data.Count + work_cell + 1).Value = min
                        work_sheet.Cells(2, data.Count + work_cell + 1).Value = max
                        work_sheet.Cells(3, data.Count + work_cell + 1).FormulaR1C1 = "AVERAGE(R[-1]C[" + (-data.Count - 1).ToString + "]:R[" + (d.Count - 3).ToString + "]C[" + (-data.Count - 1).ToString + "])"
                        work_sheet.Cells(4, data.Count + work_cell + 1).FormulaR1C1 = "STDEV(R[-2]C[" + (-data.Count - 1).ToString + "]:R[" + (d.Count - 4).ToString + "]C[" + (-data.Count - 1).ToString + "])"




                        chart.YAxis.MaxValue = CInt(max) + 1
                        chart.YAxis.MinValue = CInt(min) - 1
                        'chart.YAxis.MinValue = 0
                        chart.XAxis.CrossesAt = chart.YAxis.MinValue
                        chart.YAxis.CrossesAt = chart.XAxis.MinValue

                        work_cell = work_cell + 1

                    End If
                Next



                'remove the gridline 
                Dim nl As System.Xml.XmlNodeList = chart.ChartXml.GetElementsByTagName("c:majorGridlines")

                Dim title As System.Xml.XmlNodeList = chart.ChartXml.GetElementsByTagName("c:chartitle")

                For m As Integer = 0 To title.Count - 1
                    Dim n As System.Xml.XmlNode = nl(m)
                    Dim pn As System.Xml.XmlNode = n.ParentNode
                    pn.RemoveChild(n)
                Next


                For m As Integer = 0 To nl.Count - 1
                    Dim n As System.Xml.XmlNode = nl(m)
                    Dim pn As System.Xml.XmlNode = n.ParentNode
                    pn.RemoveChild(n)
                Next

                work_excel.Save()
            End Using
        Catch ex As Exception

        End Try

    End Sub
    '''<summary> 
    ''' Export peak find result to excel file 
    '''</summary>
    '''<param name="path">path the excel file will export to</param>
    '''<param name= "data">peak find result in the excel file</param>
    '''<remarks></remarks>
    Public Sub WriteDataToExcelForPeakFinding(ByVal path As String, ByVal data As List(Of Double()), ByVal peak_position As Double)
        Try
            Dim i As Integer = 1
            Dim j As Integer = 0

            'Dim sum As Double
            ' Dim series As Drawing.Chart.ExcelChartSerie
            Dim d As Double()
            Dim has_peak As Boolean = False
            Dim average_sensity As Double

            Dim work_cell As Integer = 1
            ' Dim work_cell_end As Integer
            Dim work_file As FileInfo = New FileInfo(path)

            Using work_excel As ExcelPackage = New ExcelPackage(work_file)


                Dim k As Integer = work_excel.Workbook.Worksheets.Count

                Dim work_sheet As ExcelWorksheet = work_excel.Workbook.Worksheets.Add(k.ToString)

                'chart. 
                While True
                    If work_sheet.Cells.Item(i, 1).Value Is Nothing = True Then
                        work_cell = i
                        Exit While
                    Else
                        i = i + 1
                    End If
                End While
                ' put x_series
                'For i = 1 To data.Item(0).Count
                '    work_sheet.Cells.Item(i, work_cell).Value = i
                'Next


                work_sheet.Cells(1, data.Count + 1).Value = "Sensity Summary"





                For Each d In data
                    Dim min As Double = Double.MaxValue
                    Dim max As Double = Double.MinValue

                    If d Is Nothing = False Then
                        j = 1
                        For Each m As Double In d
                            If (m < peak_position * 1.1 Or m > peak_position * 0.9) Then
                                has_peak = True
                            End If
                            work_sheet.Cells(j, work_cell).Value = m

                            If m < min Then min = m
                            If m > max Then max = m

                            j = j + 1

                            If j = d.Count + 1 Then
                                If has_peak Then
                                    work_sheet.Cells(j, work_cell).Value = (d.Count - 2) / (d.Count - 1)
                                    average_sensity = average_sensity + (d.Count - 2) / (d.Count - 1)
                                Else
                                    work_sheet.Cells(j, work_cell).Value = 0
                                End If

                            End If
                        Next

                        'min


                        work_cell = work_cell + 1
                        has_peak = False

                    End If
                Next


                work_sheet.Cells(2, data.Count + 2).Value = average_sensity / data.Count
                average_sensity = 0

                work_excel.Save()
            End Using
        Catch ex As Exception

        End Try

    End Sub


    Private Sub WriteDataToExcel(ByVal info As String, ByVal path As String, ByVal data As List(Of Double()), ByVal parameter() As Double, ByVal start As Integer)
        Try
            Dim i As Integer = 1
            Dim j As Integer = 0
            'Dim sum As Double
            Dim series As Drawing.Chart.ExcelChartSerie
            Dim d As Double()
            Dim first_item_index As Integer
            Dim counter As Integer

            Dim x_series As String
            Dim work_cell As Integer = 1
            ' Dim work_cell_end As Integer
            Dim work_file As FileInfo = New FileInfo(path)

            Using work_excel As ExcelPackage = New ExcelPackage(work_file)


                Dim k As Integer = work_excel.Workbook.Worksheets.Count

                Dim work_sheet As ExcelWorksheet = work_excel.Workbook.Worksheets.Add(k.ToString)
                Dim chart As Drawing.Chart.ExcelChart = work_sheet.Drawings.AddChart(k.ToString, Drawing.Chart.eChartType.XYScatterLinesNoMarkers)
                chart.Legend.Remove()
                chart.Title.Text = Nothing




                'chart. 
                While True
                    If work_sheet.Cells.Item(i, 1).Value Is Nothing = True Then
                        work_cell = i
                        Exit While
                    Else
                        i = i + 1
                    End If
                End While
                ' put x_series
                'For i = 1 To data.Item(0).Count
                '    work_sheet.Cells.Item(i, work_cell).Value = i
                'Next
                work_sheet.Cells(1, data.Count + 1).Value = "A"
                work_sheet.Cells(2, data.Count + 1).Value = "x0"
                work_sheet.Cells(3, data.Count + 1).Value = "a1"
                work_sheet.Cells(4, data.Count + 1).Value = "a2"
                work_sheet.Cells(5, data.Count + 1).Value = "B"
                work_sheet.Cells(6, data.Count + 1).Value = "C"





                For Each d In data
                    Dim min As Double = Double.MaxValue
                    Dim max As Double = Double.MinValue

                    If d Is Nothing = False Then
                        If counter Mod 2 = 0 Then
                            If d(0) > 0 Then
                                first_item_index = start + 1
                            Else
                                first_item_index = 1
                            End If
                        End If

                        j = first_item_index


                        For Each m As Double In d
                            work_sheet.Cells(j, work_cell).Value = m

                            If m < min Then min = m
                            If m > max Then max = m

                            j = j + 1
                        Next

                        If work_cell = 1 Then x_series = ExcelRange.GetAddress(1, 1, data.Item(0).Count, 1)
                        If work_cell >= 2 Then series = chart.Series.Add(ExcelRange.GetAddress(1, work_cell, data.Item(0).Count, work_cell), x_series)

                        chart.YAxis.MaxValue = CInt(max) + 1
                        chart.YAxis.MinValue = CInt(min) - 1
                        'chart.YAxis.MinValue = 0
                        chart.XAxis.CrossesAt = chart.YAxis.MinValue
                        chart.YAxis.CrossesAt = chart.XAxis.MinValue

                        work_cell = work_cell + 1

                    End If

                    counter = counter + 1
                Next

                If parameter.Length = 6 Then
                    work_sheet.Cells(1, data.Count + work_cell + 1).Value = parameter(0)
                    work_sheet.Cells(2, data.Count + work_cell + 1).Value = parameter(1)
                    work_sheet.Cells(3, data.Count + work_cell + 1).Value = parameter(2)
                    work_sheet.Cells(4, data.Count + work_cell + 1).Value = parameter(3)
                    work_sheet.Cells(5, data.Count + work_cell + 1).Value = parameter(4)
                    work_sheet.Cells(6, data.Count + work_cell + 1).Value = parameter(5)
                End If



                'remove the gridline 
                Dim nl As System.Xml.XmlNodeList = chart.ChartXml.GetElementsByTagName("c:majorGridlines")

                Dim title As System.Xml.XmlNodeList = chart.ChartXml.GetElementsByTagName("c:chartitle")

                For m As Integer = 0 To title.Count - 1
                    Dim n As System.Xml.XmlNode = nl(m)
                    Dim pn As System.Xml.XmlNode = n.ParentNode
                    pn.RemoveChild(n)
                Next


                For m As Integer = 0 To nl.Count - 1
                    Dim n As System.Xml.XmlNode = nl(m)
                    Dim pn As System.Xml.XmlNode = n.ParentNode
                    pn.RemoveChild(n)
                Next

                work_excel.Save()
            End Using
        Catch ex As Exception

        End Try

    End Sub

    Public Sub ExportFitToExcelFile(ByVal file_path As String, ByVal scale As Double, ByVal data_point_no As Integer, ByVal fit_1 As Integer, ByVal fit_2 As Integer, ByVal x0 As Double, ByVal a1 As Double, ByVal a2 As Double, ByVal c As Double, ByVal info As String, ByVal fit_data() As Double)
        Dim data As New List(Of Double())
        Dim temp_array(data_point_no - 1) As Double
        Dim temp2_array(fit_2 - fit_1) As Double
        Dim coefficients(5) As Double


        'coefficients(0) = Form1.GuessANumericEdit.Value
        coefficients(1) = x0
        coefficients(2) = a1
        ' coefficients(3) = Form1.GuessBNumericEdit.Value
        coefficients(4) = a2
        coefficients(5) = c

        For i = 0 To data_point_no - 1
            temp_array(i) = scale * i
        Next i
        data.Add(temp_array)
        data.Add(fit_data)

        For i = fit_1 To fit_2
            temp2_array(CInt(i - fit_1)) = scale * i
        Next
        data.Add(temp2_array)
        data.Add(fit_data)

        WriteDataToExcel(info, file_path + "-export" + ".xlsx", data, coefficients, fit_1)

    End Sub

    Public Sub ReadDataFromExcel(ByVal path As String)
        Dim work_file As FileInfo = New FileInfo(path)
        Using work_excel As ExcelPackage = New ExcelPackage(work_file)


        End Using
    End Sub
End Class


