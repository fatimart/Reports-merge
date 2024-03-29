﻿/**
using Microsoft.Win32;
using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static BatelcoReport.MainWindow;
using System.IO;
using Syncfusion.XlsIO;

using Color = System.Drawing.Color;
using ex = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace BatelcoReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<XlFile> XlFiles { get; } = new ObservableCollection<XlFile>();
        public ObservableCollection<Report> reports { get; } = new ObservableCollection<Report>();
        public ObservableCollection<SIMReport> BillsReports { get; } = new ObservableCollection<SIMReport>();
        public ObservableCollection<strs> str { get; } = new ObservableCollection<strs>();
        public MainWindow()
        {
            InitializeComponent();
        }
        
        public void FillDataGrid()
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "select File";
            fdlg.FileName = txtFilePath.Text;
            fdlg.DefaultExt = ".xlsx";
            fdlg.Filter = "(.xlsx)|*.xlsx;*.CSV;*.csv";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;
            fdlg.Multiselect = true;
            if (fdlg.ShowDialog() == true)
            {
                string commandText;
                string oledbConnectString;
                //txtFilePath.Text = fdlg.FileName;
                foreach (string filename in fdlg.FileNames)
                {
                    commandText = "SELECT * FROM [Sheet1$]";
                    oledbConnectString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                   @"Data Source=" + filename + ";" +
                   "Extended Properties=\"Excel 12.0;HDR=YES\";";
                    OleDbConnection connection = new OleDbConnection(oledbConnectString);
                    OleDbCommand command = new OleDbCommand(commandText, connection);
                    try
                    {
                        connection.Open();
                        DataTable dt = new DataTable();
                        OleDbDataAdapter Adpt = new OleDbDataAdapter(commandText, connection);
                        Adpt.Fill(dt);
                        dtGrid.ItemsSource = dt.DefaultView;
                    }
                    catch (Exception)
                    {
                    }
                    XlFiles.Add(new XlFile
                    {
                        Name = System.IO.Path.GetFileName(filename),
                        FullPath = filename
                    });
                }
            }

         

        }

        public void OpenMultipleFile()
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Multiselect = true;
            fdlg.Filter =
                        "(.xlsx)|*.xlsx;*.CSV;*.csv";
            fdlg.InitialDirectory =
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.MyDocuments);

            if (fdlg.ShowDialog() == true)
            {

                foreach (string filename in fdlg.FileNames)

                    XlFiles.Add(new XlFile
                    {
                        Name = System.IO.Path.GetFileName(filename),
                        FullPath = filename
                    });
            }
        }
        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenMultipleFile();

        }

        private void btnimport_Click(object sender, RoutedEventArgs e)
        {
            if(CombineWorkSheet())
            {
                ReadAll();
                Excel();
            }
            else
            {
                MessageBox.Show("Please select valid files");

            }



        }


        public bool CombineWorkSheet()
        {
            string  bills, kiosk, mPos ;
            string billsName, kioskName, mPosName;

            {
                if (XlFiles[0].Name == "bills.csv") { bills = XlFiles[0].FullPath; billsName = XlFiles[0].Name; }
                else if (XlFiles[1].Name == "bills.csv") { bills = XlFiles[1].FullPath; billsName = XlFiles[1].Name; }
                else if (XlFiles[2].Name == "bills.csv") { bills = XlFiles[2].FullPath; billsName = XlFiles[2].Name; }
                else { bills = ""; billsName = ""; }
            }
            {
                if (XlFiles[0].Name == "kiosk.xlsx") { kiosk = XlFiles[0].FullPath; kioskName = XlFiles[0].Name; }
                else if (XlFiles[1].Name == "kiosk.xlsx") { kiosk = XlFiles[1].FullPath; kioskName = XlFiles[1].Name; }
                else if (XlFiles[2].Name == "kiosk.xlsx") { kiosk = XlFiles[2].FullPath; kioskName = XlFiles[2].Name; }
                else { kiosk = ""; kioskName = ""; }
            }
            {
                if (XlFiles[0].Name == "mpos.CSV") { mPos = XlFiles[0].FullPath; mPosName = XlFiles[0].Name; }
                else if (XlFiles[1].Name == "mpos.CSV") { mPos = XlFiles[1].FullPath; mPosName = XlFiles[1].Name; }
                else if (XlFiles[2].Name == "mpos.CSV") { mPos = XlFiles[2].FullPath; mPosName = XlFiles[2].Name; }
                else { mPos = ""; mPosName = ""; }
            }

            //MessageBox.Show("the file name 1:" + bills);
            //MessageBox.Show("the file name 2:" +kiosk);
            //MessageBox.Show("the file name 3:" +mPos);

            if (billsName == "bills.csv" && kioskName == "kiosk.xlsx" && mPosName == "mpos.CSV") 
            {

                // Open Excel A file.
                Aspose.Cells.Workbook SourceBook1 = new Aspose.Cells.Workbook(bills);

                // Open Excel B file.
                Aspose.Cells.Workbook SourceBook2 = new Aspose.Cells.Workbook(kiosk);

                // Open the third excel file.
                Aspose.Cells.Workbook SourceBook3 = new Aspose.Cells.Workbook(mPos);

                // Create destination Workbook.
                Aspose.Cells.Workbook destWorkbook = new Aspose.Cells.Workbook();
                // First worksheet is added by default to the Workbook. Add the second worksheet.
                destWorkbook.Worksheets.Add();
                destWorkbook.Worksheets.Add();
                // 
                destWorkbook.Worksheets[0].Copy(SourceBook1.Worksheets[0]);

                //
                destWorkbook.Worksheets[1].Copy(SourceBook2.Worksheets[0]);
                //
                destWorkbook.Worksheets[2].Copy(SourceBook3.Worksheets[0]);

                // By default, the worksheet names are "Sheet1" and "Sheet2" respectively.
                // Lets give them meaningful names.
                destWorkbook.Worksheets[0].Name = "Sheet1";
                destWorkbook.Worksheets[1].Name = "Sheet2";
                destWorkbook.Worksheets[2].Name = "Sheet3";

                // Save the destination file.
                destWorkbook.Save("CombinedFile1.xlsx");

                return true;
            }

            else
            {
                MessageBox.Show("Please select 3 files with names bills, kiosk and mpos");
                XlFiles.Clear();
                return false;
            }

        }
        //public void combineMultiWorkSheet()
        //{
        //    // Open an Excel file that contains the worksheets:
        //    // Products1, Products2 and Products3
        //    Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("CombinedFile1.xlsx");

        //    // Add a worksheet named Summary_sheet
        //    Aspose.Cells.Worksheet summarySheet = workbook.Worksheets.Add("Summary_sheet");

        //    // Iterate over source worksheets whose data you want to copy to the
        //    // summary worksheet
        //    string[] nameOfSourceWorksheets = { "Sheet1", "Sheet2", "Sheet3" };
        //    int totalRowCount = 0;

        //    foreach (string sheetName in nameOfSourceWorksheets)
        //    {
        //        Aspose.Cells.Worksheet sourceSheet = workbook.Worksheets[sheetName];

        //        Range sourceRange;
        //        Range destRange;
        //        // In case of Sheet1 worksheet, include all rows and cols.
        //        if (sheetName.Equals("Sheet1"))
        //        {
        //            sourceRange = sourceSheet.Cells.MaxDisplayRange;

        //            destRange = summarySheet.Cells.CreateRange(
        //                    sourceRange.FirstRow + totalRowCount,
        //                    sourceRange.FirstColumn,
        //                    sourceRange.RowCount,
        //                    sourceRange.ColumnCount);
        //        }
        //        // In case of Products2 and Products3 worksheets,
        //        // exclude the first row (which contains headings).
        //        else
        //        {
        //            int mdatarow = sourceSheet.Cells.MaxDataRow; // Zero-based
        //            int mdatacol = sourceSheet.Cells.MaxDataColumn; // Zero-based
        //            sourceRange = sourceSheet.Cells.CreateRange(0 + 1, 0, mdatarow, mdatacol + 1);

        //            destRange = summarySheet.Cells.CreateRange(
        //                    sourceRange.FirstRow + totalRowCount - 1,
        //                    sourceRange.FirstColumn,
        //                    sourceRange.RowCount,
        //                    sourceRange.ColumnCount);
        //        }

        //        // Copies data, formatting, drawing objects etc. from a
        //        // source range to destination range.
        //        destRange.Copy(sourceRange);
        //        totalRowCount = sourceRange.RowCount + totalRowCount;
        //    }

        //    // Save the workbook 
        //    workbook.Save("Summarized.xlsx");

        //}

        public void read1 ()
        {
            //  Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("CombinedFile1.xlsx");
            string CombineFile = "CombinedFile1.xlsx";
            string commandText = "SELECT * FROM [Sheet1$]";
            string oledbConnectString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            @"Data Source=" + CombineFile + ";" +
            "Extended Properties=\"Excel 12.0;HDR=YES\";";
            OleDbConnection connection = new OleDbConnection(oledbConnectString);
            OleDbCommand command = new OleDbCommand(commandText, connection);
            //   DataTable dt = new DataTable();
            // OleDbDataAdapter Adpt = new OleDbDataAdapter(commandText, connection);
            OleDbDataReader reader;
            try
            {
                connection.Open();
                reader = command.ExecuteReader();
                //  Adpt.Fill(dt);
                //   dtGrid.ItemsSource = dt.DefaultView;
                while (reader.Read())
                {
                    if (reader["Transaction Status "].ToString().Trim() == "SUCCESS")
                    {
                        BillsReports.Add(new SIMReport
                        {
                            ACCOUNT_NUMBER = (reader["Customer Phone Number "]).ToString(),
                            CUSTOMER_NAME = reader["First Name "].ToString().Trim() + " " + reader["Last Name "].ToString().Trim(),
                            TRANSACTION_NUMBER = reader["Transaction Id "].ToString(),

                            PAYMENTDATE = Convert.ToDateTime(reader["Transaction Date "]),

                            Date_of_payment_execution = default,
                            PRODUCTAMOUNT = Convert.ToDouble(reader["Transaction Amount "]),

                            Commission = default,
                            VATAMOUNT = default,
                            Net_Amount = Convert.ToDouble(reader["Transaction Amount "]),
                            KIOSKID = default,

                            PRODUCTNAME = "SIM",

                            ORDERNUMBER = reader["Reference Number Provider "].ToString().Trim(),

                            PAYMENTLOCATION = "SIM",

                            Transaction_Status = reader["Transaction Status "].ToString().Trim(),
                        });
                    }

                }


                connection.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message);
                connection.Close();
            }


        }
        public void read2()
        {
            //  Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("CombinedFile1.xlsx");
            string CombineFile = "CombinedFile1.xlsx";
            string commandText = "SELECT * FROM [Sheet2$]";
            string oledbConnectString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            @"Data Source=" + CombineFile + ";" +
            "Extended Properties=\"Excel 12.0;HDR=YES\";";
            OleDbConnection connection = new OleDbConnection(oledbConnectString);
            OleDbCommand command = new OleDbCommand(commandText, connection);
            DataTable dt = new DataTable();
        //    OleDbDataAdapter Adpt = new OleDbDataAdapter(commandText, connection);
            OleDbDataReader reader;
            try
            {
                connection.Open();
                reader = command.ExecuteReader();
          //      Adpt.Fill(dt);
                //    dtGrid.ItemsSource = dt.DefaultView;
                while (reader.Read())
                {
                    if (reader["BT Res"].ToString() == "Success")
                    {
                        reports.Add(new Report
                        {
                            ACCOUNT_NUMBER = Convert.ToInt32(reader["Phone Number"]),
                            CUSTOMER_NAME = default,
                            TRANSACTION_NUMBER = reader["Batelco Transaction ID"].ToString(),

                            PAYMENTDATE = Convert.ToDateTime(reader["Dateof Payment Received"]),

                            Date_of_payment_execution = default,
                            AMOUNT = Convert.ToDouble(reader["Trx Amount"]),

                            Commission = Convert.ToDouble(reader["Commission"]),
                            VAT = default,
                            Net_Amount = Convert.ToDouble(reader["Net Amount"]),
                            AUTHRIZATION_NO = default,

                            Service_Name = reader["Channel Name"].ToString(),

                            REFERENCE_NO = reader["YQ Transaction ID"].ToString(),

                            PAYMENTLOCATION = default,

                            Transaction_Status = reader["BT Res"].ToString(),
                        });
                    }
                    //connection.Open();
                    //int rowsAffected = command.ExecuteNonQuery();
                    //connection.Close();
                }


                connection.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message);
                connection.Close();
            }
        }
        public void read3()
        {
            string CombineFile = "CombinedFile1.xlsx";
            string commandText = "SELECT * FROM [Sheet3$]";
            string oledbConnectString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            @"Data Source=" + CombineFile + ";" +
            "Extended Properties=\"Excel 12.0;HDR=YES\";";
            OleDbConnection connection = new OleDbConnection(oledbConnectString);
            OleDbCommand command = new OleDbCommand(commandText, connection);
            DataTable dt = new DataTable();
            OleDbDataReader reader;

            try
            {
                connection.Open();
                reader = command.ExecuteReader();
             //   Adpt.Fill(dt);
                //  dtGrid.ItemsSource = dt.DefaultView;
                while (reader.Read())
                {
                    if (reader["Transaction Status"].ToString().Trim()== "Success") {
                        reports.Add(new Report
                        {
                            ACCOUNT_NUMBER = Convert.ToInt32(reader["Circuit Number"]),
                            CUSTOMER_NAME = default,
                            TRANSACTION_NUMBER = reader["Batelco Transaction ID"].ToString(),

                            PAYMENTDATE = Convert.ToDateTime(reader["Date of Payment Recieved"]),

                            Date_of_payment_execution = default,
                            AMOUNT = Convert.ToDouble(reader["Trx Amount"]),

                            Commission = Convert.ToDouble(reader["Commission"]),
                            VAT = default,
                            Net_Amount = Convert.ToDouble(reader["Net Amount"]),
                            AUTHRIZATION_NO = default,

                            Service_Name = reader["Channel Name"].ToString(),

                            REFERENCE_NO = reader["YQ Transactio ID"].ToString(),

                            PAYMENTLOCATION = default,

                            Transaction_Status = reader["Transaction Status"].ToString(),
                        });
                    }
                    
                }


                connection.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message);
                connection.Close();
            }
        }

        public void ReadAll()
        {
             read1();
             read2();
             read3();
        }



        public void Excel ()
        {

            Microsoft.Office.Interop.Excel.Application excelApp;
            // load excel, and create a new workbook
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            ex.Workbooks workb = excelApp.Workbooks;
            //excelApp.Visible = true;

            workb.Add();

            // single worksheet
            Microsoft.Office.Interop.Excel.Worksheet worksheet = excelApp.ActiveSheet;
            worksheet.Name = "YQ";
            excelApp.ActiveWindow.DisplayGridlines = false;
            int c = reports.Count - 2;




            //cells size
            worksheet.Range["A1"].ColumnWidth = 21.43;
            worksheet.Range["B1"].ColumnWidth = 45.14;
            worksheet.Range["C1"].ColumnWidth = 22.43;
            worksheet.Range["D1"].ColumnWidth = 25.71;
            worksheet.Range["E1"].ColumnWidth = 24.71;
            worksheet.Range["F1"].ColumnWidth = 17.57;
            worksheet.Range["G1"].ColumnWidth = 11.14;
            worksheet.Range["H1"].ColumnWidth = 12.43;
            worksheet.Range["I1"].ColumnWidth = 27.57;
            worksheet.Range["J1"].ColumnWidth = 17.71;
            worksheet.Range["K1"].ColumnWidth = 15.43;
            worksheet.Range["L1"].ColumnWidth = 39.29;
            worksheet.Range["M1"].ColumnWidth = 18.57;
            worksheet.Range["N1"].ColumnWidth = 16.57;

            //style and element
            worksheet.Cells[3, 1] = "ACCOUNT NUMBER";
            worksheet.Cells[3, 2] = "CUSTOMER NAME";
            worksheet.Cells[3, 3] = "TRANSACTION NUMBER";
            worksheet.Cells[3, 4] = "PAYMENTDATE";
            worksheet.Cells[3, 5] = "Date of payment execution";
            worksheet.Cells[3, 6] = "AMOUNT";
            worksheet.Cells[3, 7] = "Commission";
            worksheet.Cells[3, 8] = "VAT";
            worksheet.Cells[3, 9] = "Net_Amount";
            worksheet.Cells[3, 10] = "AUTHRIZATION_NO";
            worksheet.Cells[3, 11] = "Service_Name";
            worksheet.Cells[3, 12] = "REFERENCE_NO";
            worksheet.Cells[3, 13] = "PAYMENTLOCATION";
            worksheet.Cells[3, 14] = "Transaction_Status";

            worksheet.Cells[2, 1] = "YQ (Kiosk, MPOS, TAM)";
            worksheet.Cells[2, 4] = "TrnasactionDate";
            worksheet.Cells[2, 5] = "To bank Account";
            worksheet.Cells[2, 9] = "Transferred to Account";
            worksheet.Cells[2, 10] = "To be empty";
            worksheet.Cells[2, 13] = "To be defaultesd";
            worksheet.Cells[2, 14] = "Success only";

            worksheet.get_Range("A2:N2").Font.Bold = true;
            worksheet.get_Range("A3:N3").Font.Bold = true;
            worksheet.get_Range("A3", "N3").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.get_Range("A3", "N3").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            worksheet.get_Range("A4", "N" + c + "").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            worksheet.get_Range("A4", "N" + c + "").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;

            worksheet.get_Range("C3", "C" + c + "").Interior.Color = System.Drawing.Color.Yellow;
            worksheet.get_Range("L3", "L" + c + "").Interior.Color = System.Drawing.Color.Yellow;

            worksheet.get_Range("A3", "N" + c + "").Borders.LineStyle = ex.XlLineStyle.xlContinuous;
            worksheet.get_Range("A3", "N" + c + "").Borders.Weight = ex.XlBorderWeight.xlThin;

            ////Text Alignment Setting (Horizontal Alignment)
            // worksheet.get_Range("A3:N3").HorizontalAlignment = ExcelHAlign.HAlignCenter;

            ////Color
            //worksheet.get_Range("A2:N2").Font.Color = Color.FromArgb(255, 0, 0);
            //worksheet.Cells["C3"].CellStyle.Color = Color.FromArgb(255, 255, 0);
            //worksheet.Cells["L3"].CellStyle.Color = Color.FromArgb(255, 255, 0);

            Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;
            Microsoft.Office.Interop.Excel.Range cell = range.Cells[1, 14];
            Microsoft.Office.Interop.Excel.Borders border = cell.Borders;



            //border[ex.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            //border[ex.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            //border[ex.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            //border[ex.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


            for (int i = 4, n = 0; i < reports.Count - 1; i++, n++)
            {


                worksheet.Cells[i, 1].Value = reports[n].ACCOUNT_NUMBER;
                worksheet.Cells[i, 2].Value = reports[n].CUSTOMER_NAME;
                worksheet.Cells[i, 3].Value = reports[n].TRANSACTION_NUMBER;
                //worksheet.Cells[i, 3].CellStyle.Color = Color.FromArgb(255, 255, 0);
                worksheet.Cells[i, 4].Value = reports[n].PAYMENTDATE.ToString("dd / MM / yyyy HH: mm:ss");
                //worksheet.Cells[i, 4].NumberFormat = "dd/mm/yyyy hh:mm:ss";

                worksheet.Cells[i, 5].Value = reports[n].Date_of_payment_execution;
                worksheet.Cells[i, 6].Value = reports[n].AMOUNT;
                //worksheet.Cells[i, 6].Value = reports[n].AMOUNT.ToString(" 0.000");
                //worksheet.Cells[i, 6].NumberFormat = "0.000";

                worksheet.Cells[i, 7].Value = reports[n].Commission;
                //worksheet.Cells[i, j].NumberFormat = "0.000";

                worksheet.Cells[i, 8].Value = reports[n].VAT;
                //worksheet.Cells[i, j].NumberFormat = "0.000";

                worksheet.Cells[i, 9].Value = (reports[n].AMOUNT - reports[n].Commission - reports[n].VAT);
                //worksheet.Cells[i, j].NumberFormat = "0.000";

                worksheet.Cells[i, 10].Value = reports[n].AUTHRIZATION_NO;
                worksheet.Cells[i, 11].Value = reports[n].Service_Name;
                worksheet.Cells[i, 12].Value = reports[n].REFERENCE_NO;
                //worksheet.Cells[i, 12].CellStyle.Color = Color.FromArgb(255, 255, 0);

                worksheet.Cells[i, 13].Value = reports[n].PAYMENTLOCATION;
                worksheet.Cells[i, 14].Value = reports[n].Transaction_Status;

                //workSheet_range.Interior.Color = System.Drawing.Color.Yellow.ToArgb();

            }



            int j = reports.Count + 1;


            worksheet.Cells[j, 1] = "YQ (SIM)";
            worksheet.Cells[j, 4] = "TrnasactionDate";
            worksheet.Cells[j, 5] = "To bank Account";
            worksheet.Cells[j, 9] = "Transferred to Account";
            worksheet.Cells[j, 10] = "To be empty";
            worksheet.Cells[j, 13] = "To be defaultesd";
            worksheet.Cells[j, 14] = "Success only";
            j++;

            worksheet.Cells[j, 1] = "ACCOUNT NUMBER";
            worksheet.Cells[j, 2] = "CUSTOMER NAME";
            worksheet.Cells[j, 3] = "TRANSACTION_NUMBER";
            worksheet.Cells[j, 4] = "PAYMENTDATE";
            worksheet.Cells[j, 5] = "Date of payment execution";
            worksheet.Cells[j, 6] = "PRODUCTAMOUNT";
            worksheet.Cells[j, 7] = "Commission";
            worksheet.Cells[j, 8] = "VATAMOUNT";
            worksheet.Cells[j, 9] = "TOTALTRANSACTIONAMOUNT";
            worksheet.Cells[j, 10] = "KIOSKID";
            worksheet.Cells[j, 11] = "PRODUCTNAME";
            worksheet.Cells[j, 12] = "ORDERNUMBER";
            worksheet.Cells[j, 13] = "PAYMENTLOCATION";
            worksheet.Cells[j, 14] = "Transaction Status";

            for (int i = j; i <= 14; i++)
            {
                worksheet.get_Range("A'" + j + "':N'" + j + "'").Font.Bold = true;

            }

            j++;

            for (int i = j, n = 0; n < BillsReports.Count - 1; n++, i++)
            {
                worksheet.Cells[i, 1].Value = BillsReports[n].ACCOUNT_NUMBER;
                worksheet.Cells[i, 2].Value = BillsReports[n].CUSTOMER_NAME;
                worksheet.Cells[i, 3].Value = BillsReports[n].TRANSACTION_NUMBER;
                worksheet.Cells[i, 4].Value = BillsReports[n].PAYMENTDATE.ToString("dd / MM / yyyy HH: mm:ss");
                worksheet.Cells[i, 5].Value = BillsReports[n].Date_of_payment_execution;
                worksheet.Cells[i, 6].Value = BillsReports[n].PRODUCTAMOUNT;
                worksheet.Cells[i, 7].Value = BillsReports[n].Commission;
                worksheet.Cells[i, 8].Value = BillsReports[n].VATAMOUNT;
                worksheet.Cells[i, 9].Value = (BillsReports[n].PRODUCTAMOUNT - BillsReports[n].Commission - BillsReports[n].VATAMOUNT);
                worksheet.Cells[i, 10].Value = BillsReports[n].KIOSKID;
                worksheet.Cells[i, 11].Value = BillsReports[n].PRODUCTNAME;
                worksheet.Cells[i, 12].Value = BillsReports[n].ORDERNUMBER;
                worksheet.Cells[i, 13].Value = BillsReports[n].PAYMENTLOCATION;
                worksheet.Cells[i, 14].Value = BillsReports[n].Transaction_Status;


            }
            int C1 = j + BillsReports.Count - 2;
            worksheet.get_Range("A" + (j - 1) + "", "N" + (j - 1) + "").Font.Bold = true;
            worksheet.get_Range("A" + (j - 1) + "", "N" + (j - 1) + "").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.get_Range("A" + (j - 1) + "", "N" + (j - 1) + "").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            worksheet.get_Range("A" + j + "", "N" + C1 + "").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            worksheet.get_Range("A" + j + "", "N" + C1 + "").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;

            worksheet.get_Range("A" + (j - 1) + "", "N" + C1 + "").Borders.LineStyle = ex.XlLineStyle.xlContinuous;
            worksheet.get_Range("A" + (j - 1) + "", "N" + C1 + "").Borders.Weight = ex.XlBorderWeight.xlThin;
            worksheet.Cells[(C1 + 2), 4].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Save Excel sheet";
            saveFileDialog1.Filter = "Excel files|*.xlsx|All files|*.*";
            saveFileDialog1.FileName = "YQ_Sample_Report" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx";


            if (saveFileDialog1.ShowDialog() == true)
            {
                //Get the FileInfo
                FileInfo fi = new FileInfo(saveFileDialog1.FileName);

                //write the file to the disk
                worksheet.SaveAs(saveFileDialog1.FileName,
                                 ex.XlFileFormat.xlOpenXMLWorkbook,
                                 Type.Missing,
                                 Type.Missing,
                                 false,
                                 Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
                                 false,
                                 false,
                                 Type.Missing,
                                 Type.Missing);
            }

            this.Close();
            excelApp.Quit();
            System.Diagnostics.Process.Start(saveFileDialog1.FileName);

        }
        public class strs
        {
            public int ACCOUNT_NUMBER { get; set; }
            public string CUSTOMER_NAME { get; set; }

        }

     
    }
}




**/