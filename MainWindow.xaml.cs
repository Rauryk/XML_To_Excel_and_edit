using DocumentFormat.OpenXml.EMMA;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;
using System.Data;
using OfficeExcel = Microsoft.Office.Interop.Excel;
namespace Militaria_Gruszecki_Roman
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }



        private async void Open_File_Button_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                try
                {
                    DataSet ds = new DataSet();
                    ds.ReadXml(openFileDialog.FileName);
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        DataTable dt = new DataTable("arkusz1");
                        ds.Tables[0].Columns.Add("ilosc_zdjęć");
                        ds.Tables[0].Columns.Add("marża"); 
                        DataTable dt2 = new DataTable("arkusz2");
                        ds.Tables[2].Columns.Remove(ds.Tables[2].Columns[0]);
                        foreach (DataColumn column in ds.Tables[0].Columns)
                        {
                            dt.Columns.Add(column.ColumnName);

                        }

                        DataRow dr = dt.NewRow();
                        foreach (DataColumn column in ds.Tables[0].Columns)
                        {
                            dr[column.ColumnName] = column.ColumnName;

                        }
                        dt.Rows.Add(dr);

                        foreach (DataColumn column in ds.Tables[2].Columns)
                        {
                            dt2.Columns.Add(column.ColumnName);

                        }

                        DataRow dr2 = dt2.NewRow();
                        foreach (DataColumn column in ds.Tables[2].Columns)
                        {
                            dr2[column.ColumnName] = column.ColumnName;

                        }
                        dt2.Rows.Add(dr2);
                        List<int> ints = new List<int>();
                        foreach (DataRow row in ds.Tables[2].Rows)
                        {
                            ints.Add(Int32.Parse(row[1].ToString()));
                            dt2.ImportRow(row);

                        }
                        int int_row = 0;
                        foreach (DataRow row in ds.Tables[0].Rows)
                        {
                            row[14] = ints.Where(x => x == int_row).Count();
                            row[15] = (decimal.Parse(row[11].ToString()) - decimal.Parse(row[10].ToString())) / decimal.Parse(row[10].ToString()) * 100;
                            int_row++;
                            dt.ImportRow(row);

                        }

                        var ws = wb.Worksheets.Add(dt.TableName);
                        ws.Cell(1, 1).InsertData(dt.Rows);

                        var ws2 = wb.Worksheets.Add(dt2.TableName);
                        ws2.Cell(1, 1).InsertData(dt2.Rows);
                        ws2.Columns().AdjustToContents();
                        int skok = 1;
                        foreach (var row in ws.Rows())
                        {
                            if (skok == 1)
                            {
                                skok = 0;
                                continue;

                            }

                            if (decimal.Parse(row.Cell(16).Value.ToString()) <= 25)
                            {

                                row.Style.Fill.BackgroundColor = XLColor.Red;
                            }
                            if (decimal.Parse(row.Cell(15).Value.ToString()) <= 2)
                            {

                                row.Style.Fill.BackgroundColor = XLColor.Orange;
                            }

                        }
                        ws.Columns().AdjustToContents();


                        wb.SaveAs("basa_xml.xlsx");
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();

                    }


                }
                catch (Exception ex)
                {

                    Console.WriteLine(ex.GetBaseException());
                }


        }
        public System.Data.DataTable formofDataTable(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            string worksheetName = ws.Name;
            dt.TableName = worksheetName;
            Microsoft.Office.Interop.Excel.Range xlRange = ws.UsedRange;
            object[,] valueArray = (object[,])xlRange.get_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault);
            for (int k = 1; k <= valueArray.GetLength(1); k++)
            {
                dt.Columns.Add((string)valueArray[1, k]);  
            }
            object[] singleDValue = new object[valueArray.GetLength(1)]; 
            for (int i = 2; i <= valueArray.GetLength(0); i++)
            {
                for (int j = 0; j < valueArray.GetLength(1); j++)
                {
                    if (valueArray[i, j + 1] != null)
                    {
                        singleDValue[j] = valueArray[i, j + 1].ToString();
                    }
                    else
                    {
                        singleDValue[j] = valueArray[i, j + 1];
                    }
                }
                dt.LoadDataRow(singleDValue, System.Data.LoadOption.PreserveChanges);
            }

            return dt;
        }
        DataSet ds = new DataSet();
        private void open_xlsx_button_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog choofdlog = new OpenFileDialog();
            if (choofdlog.ShowDialog() == true)
            {
                string sFileName = choofdlog.FileName;
                string path = System.IO.Path.GetFullPath(choofdlog.FileName);
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                ds.Reset();
                Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(path);
                foreach (Microsoft.Office.Interop.Excel.Worksheet ws in wb.Worksheets)
                {
                    System.Data.DataTable td = new System.Data.DataTable();
                    td = formofDataTable(ws);
                    ds.Tables.Add(td);
                }

                data_excel.ItemsSource = ds.Tables[0].AsDataView();
                data_excel2.ItemsSource = ds.Tables[1].AsDataView();
                wb.Close();
            }


        }

        private void ExportDataSetToExcel(DataSet ds, string strPath)
        {
            int  inColumn = 0, inRow = 0;
            System.Reflection.Missing Default = System.Reflection.Missing.Value;
            //Create File
            strPath += @"\basa_xml_edit.xlsx";
            OfficeExcel.Application excelApp = new OfficeExcel.Application();
            OfficeExcel.Workbook excelWorkBook = excelApp.Workbooks.Add(1);
            foreach (DataTable dtbl in ds.Tables)
            {
                //Create Excel WorkSheet
                OfficeExcel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add(Default, excelWorkBook.Sheets[excelWorkBook.Sheets.Count], 1, Default);

                //Column Name
                for (int i = 0; i < dtbl.Columns.Count; i++)
                    excelWorkSheet.Cells[ 1, i + 1] = dtbl.Columns[i].ColumnName.ToUpper();

                // Rows
                for (int m = 0; m < dtbl.Rows.Count; m++)
                {
                    
                    for (int n = 0; n < dtbl.Columns.Count; n++)
                    {
                        inColumn = n + 1;
                        inRow =  2 + m;
                        excelWorkSheet.Cells[inRow, inColumn] = dtbl.Rows[m].ItemArray[n].ToString();
                        if (n==15 && decimal.Parse(dtbl.Rows[m].ItemArray[15].ToString()) <= 25)
                        {
                            excelWorkSheet.get_Range("A" + inRow.ToString(), "P" + inRow.ToString()).Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ff0000");
                        }
                        if (n==14 && decimal.Parse(dtbl.Rows[m].ItemArray[14].ToString()) <= 2)
                        {
                            excelWorkSheet.get_Range("A" + inRow.ToString(), "P" + inRow.ToString()).Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ff9933");
                        }
                    }
                }

            }

            //Delete First Page
            excelApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Worksheet lastWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[1];
            lastWorkSheet.Delete();
            excelApp.DisplayAlerts = true;
          
            //Set Defualt Page
            (excelWorkBook.Sheets[1] as OfficeExcel._Worksheet).Activate();
          
           excelWorkBook.SaveAs(strPath, Default, Default, Default, false, Default, OfficeExcel.XlSaveAsAccessMode.xlNoChange, Default, Default, Default, Default, Default);
           excelWorkBook.Close();
           excelApp.Quit();
        }
        private void Save_button_Click(object sender, RoutedEventArgs e)
        {
            ExportDataSetToExcel(ds.Tables[0].DataSet, System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName));
        }
        bool x = false;
        private void change_button_Click(object sender, RoutedEventArgs e)
        {

            if (x)
            {
                data_excel.Visibility = Visibility.Visible;
                data_excel2.Visibility = Visibility.Hidden;
                x = !x; 
            }
            else
            {
                data_excel.Visibility = Visibility.Hidden;
                data_excel2.Visibility = Visibility.Visible;
                x = !x;
            }            
        }
    }
}
