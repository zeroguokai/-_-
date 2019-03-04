Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports System.Threading.Tasks
Imports System.Text

Class MainWindow
    Private 状态 As Boolean = True
    Private 汇总表路径 As String
    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Dim a As New OpenFileDialog
        Dim s As New StringBuilder
        a.Filter = "Excel(*.xls;*.xlsx)|*.xls;*.xlsx"
        a.ShowDialog()
        汇总表路径文本框.Text = a.FileName
        汇总表路径 = a.FileName.Substring(0, a.FileName.Count - a.SafeFileName.Count)
    End Sub

    Private Sub Button_Click_2(sender As Object, e As RoutedEventArgs)
        Dim a As New FolderBrowserDialog
        a.ShowDialog()
        清单表文件夹路径文本框.Text = a.SelectedPath
    End Sub

    Private Async Sub kaishi(sender As Object, e As RoutedEventArgs)
        Dim exc As New Excel.Application
        Dim 汇总表 As Excel.Workbook
        Dim 汇总表s As Excel.Worksheet
        Dim 行号 As Integer = 1
        Dim excelfile As FileInfo()
        Dim 故障列表 As New List(Of String)
        清单表文件夹路径文本框.IsEnabled = False
        汇总表路径文本框.IsEnabled = False
        B1.IsEnabled = False
        B2.IsEnabled = False
        B3.IsEnabled = False
        Try
            汇总表 = exc.Workbooks.Open(汇总表路径文本框.Text.ToString)
            汇总表s = 汇总表.Worksheets("铁塔站点清查")
        Catch ex As Exception
            MessageBox.Show("汇总表路径设置错误")
            exc.Quit()
            Return
        End Try
        Try
            Dim ds As New DirectoryInfo(清单表文件夹路径文本框.Text.ToString)
            excelfile = ds.GetFiles("*.xls?", SearchOption.AllDirectories)
        Catch ex As Exception
            MessageBox.Show("清单表文件夹路径设置错误")
            exc.Quit()
            Return
        End Try
        汇总表s.UsedRange.Clear()
        Try
            汇总表.Worksheets("表头").Range(汇总表.Worksheets("表头").Cells(1, 1), 汇总表.Worksheets("表头").Cells(2, 63)).Copy()
            汇总表s.Range(汇总表s.Cells(1, 1), 汇总表s.Cells(2, 63)).PasteSpecial()
        Catch ex As Exception
            MessageBox.Show("汇总表中工作表""表头""损坏")
            exc.Quit()
            Return
        End Try
        For Each wb In excelfile
            状态框.Text = String.Format("进度：     成功导入数：{0} / 导入错误数：{1} / 总数：{2} / 百分比进度：{3}%", 行号 - 1, 故障列表.Count, excelfile.Count， Math.Round(（行号 - 1 + 故障列表.Count） / excelfile.Count, 2) * 100)
            Await ChuLi(wb.FullName, 汇总表s, exc, 行号)
            If 状态 = False Then
                故障列表.Add(wb.FullName)
            Else
                行号 = 行号 + 1
            End If
        Next
        If 故障列表.Count <> 0 Then
            Dim txt = File.CreateText(String.Concat(汇总表路径, "未正确导入的清单.txt"))
            For Each gz In 故障列表
                txt.WriteLine(gz)
                txt.Flush()
            Next
            txt.Close()
            状态框.Text = String.Format("导入完成。  完成度：{0} / {1}。   发现{2}张清单表导入错误，详见文件{3}", 行号 - 1, excelfile.Count, 故障列表.Count, String.Concat(汇总表路径, "未正确导入的清单.txt"))
        Else
            状态框.Text = String.Format("导入完成。  完成度：{0} / {1}。   全部清单已导入，未发现错误", 行号 - 1, excelfile.Count)
        End If
        exc.Visible = True
        清单表文件夹路径文本框.IsEnabled = True
        汇总表路径文本框.IsEnabled = True
        B1.IsEnabled = True
        B2.IsEnabled = True
        B3.IsEnabled = True
    End Sub
    Private Function ChuLi(wb As String, wbhz As Excel.Worksheet, exc As Excel.Application, 行号 As Integer) As Task
        Return Task.Run（Sub()
                            Dim 清单表 As Excel._Workbook
                            Try
                                清单表 = exc.Workbooks.Open(wb)
                                清单表.Worksheets.Item("模板").Range(清单表.Worksheets("模板").Cells(4, 3), 清单表.Worksheets("模板").Cells(61, 3)).Copy
                                wbhz.Range(wbhz.Cells.Item(2 + 行号, 4), wbhz.Cells.Item(2 + 行号, 61)).PasteSpecial(Transpose:=True, Paste:=Excel.XlPasteType.xlPasteValues)
                                清单表.Worksheets.Item("模板").Cells(2, 1).Copy
                                wbhz.Cells.Item(2 + 行号, 3).PasteSpecial
                                清单表.Close()
                                状态 = True
                            Catch ex As Exception
                                Try
                                    清单表.Close()
                                    状态 = False
                                Catch ex1 As Exception
                                    状态 = False
                                End Try
                            End Try
                        End Sub)

    End Function
End Class

