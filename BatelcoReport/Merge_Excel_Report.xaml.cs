using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.Windows;

using static BatelcoReport.MainWindow;
using System.IO;

using Color = System.Drawing.Color;
using ex = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace BatelcoReport
{
    /// <summary>
    /// Interaction logic for Merge_Excel_Report.xaml
    /// </summary>
    public partial class Merge_Excel_Report : Window
    {
        public ObservableCollection<XlFile> XlFiles { get; } = new ObservableCollection<XlFile>();
        public ObservableCollection<Report> reports { get; } = new ObservableCollection<Report>();
        public ObservableCollection<SIMReport> BillsReports { get; } = new ObservableCollection<SIMReport>();
        public ObservableCollection<strs> str { get; } = new ObservableCollection<strs>();


        public Merge_Excel_Report ()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click ( object sender, RoutedEventArgs e )
        {

            OpenMultipleFile();


        }

        private void btnExport_Click ( object sender, RoutedEventArgs e )
        {
            if (CombineWorkSheet())
            {
                ReadAll();
                Excel();
            }
            
        }

        private void DeleteButton_Click ( object sender, RoutedEventArgs e )

        {
            try
            {

      
            if (listboxFiles.Items.Count > 0)
            {
                listboxFiles.Items.RemoveAt

                    (listboxFiles.Items.IndexOf(listboxFiles.SelectedItem));
            }
            }
            catch 
            {

             
            }
        }
        public void OpenMultipleFile ()
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Multiselect = true;
            fdlg.Filter = "(.xlsx)|*.xlsx;*.CSV;*.csv";
            fdlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (fdlg.ShowDialog() == true)
            {

                foreach (string filename in fdlg.FileNames)
                {
                    if (System.IO.Path.GetFileName(filename) == "bills.csv" || System.IO.Path.GetFileName(filename) == "kiosk.xlsx" || System.IO.Path.GetFileName(filename) == "mpos.CSV" && listboxFiles.Items.Count <= 3)
                    {
                        XlFiles.Add(new XlFile
                        {
                            Name = System.IO.Path.GetFileName(filename),
                            FullPath = filename
                        });


                        listboxFiles.Items.Add(System.IO.Path.GetFullPath(filename));
                       
                    }
                    else
                    {
                            MessageBox.Show("You Are Allowed to select 3 files only");
                        
                    }
                }

            }

        }

        public bool CombineWorkSheet ()
        {
            

            if (listboxFiles.Items.Count != 0)
            {
                string bills, kiosk, mPos;
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
            else
            {
                MessageBox.Show("Please select 3 files with names bills, kiosk and mpos");
                XlFiles.Clear();
                return false;
            }
        }

        public void read1 ()
        {
            string CombineFile = "CombinedFile1.xlsx";
            string commandText = "SELECT * FROM [Sheet1$]";
            string oledbConnectString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            @"Data Source=" + CombineFile + ";" +
            "Extended Properties=\"Excel 12.0;HDR=YES\";";
            OleDbConnection connection = new OleDbConnection(oledbConnectString);
            OleDbCommand command = new OleDbCommand(commandText, connection);
            
            OleDbDataReader reader;
            try
            {
                connection.Open();
                reader = command.ExecuteReader();
               
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

        public void read2 ()
        {
            string CombineFile = "CombinedFile1.xlsx";
            string commandText = "SELECT * FROM [Sheet2$]";
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
                    
                }


                connection.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message);
                connection.Close();
            }
        }
        public void read3 ()
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
               
                while (reader.Read())
                {
                    if (reader["Transaction Status"].ToString().Trim() == "Success")
                    {
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

        public void ReadAll ()
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

            worksheet.get_Range("A2:N2").Font.Color = Color.FromArgb(255, 0, 0);
            worksheet.get_Range("C3", "C" + c + "").Interior.Color = System.Drawing.Color.Yellow;
            worksheet.get_Range("L3", "L" + c + "").Interior.Color = System.Drawing.Color.Yellow;

            worksheet.get_Range("A3", "N" + c + "").Borders.LineStyle = ex.XlLineStyle.xlContinuous;
            worksheet.get_Range("A3", "N" + c + "").Borders.Weight = ex.XlBorderWeight.xlThin;


            Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;
            Microsoft.Office.Interop.Excel.Range cell = range.Cells[1, 14];
            Microsoft.Office.Interop.Excel.Borders border = cell.Borders;




            for (int i = 4, n = 0; i < reports.Count - 1; i++, n++)
            {


                worksheet.Cells[i, 1].Value = reports[n].ACCOUNT_NUMBER;
                worksheet.Cells[i, 2].Value = reports[n].CUSTOMER_NAME;
                worksheet.Cells[i, 3].Value = reports[n].TRANSACTION_NUMBER;
                worksheet.Cells[i, 4].Value = reports[n].PAYMENTDATE.ToString("dd / MM / yyyy HH: mm:ss");

                worksheet.Cells[i, 5].Value = reports[n].Date_of_payment_execution;
                worksheet.Cells[i, 6].Value = reports[n].AMOUNT;

                worksheet.Cells[i, 7].Value = reports[n].Commission;

                worksheet.Cells[i, 8].Value = reports[n].VAT;

                worksheet.Cells[i, 9].Value = (reports[n].AMOUNT - reports[n].Commission - reports[n].VAT);

                worksheet.Cells[i, 10].Value = reports[n].AUTHRIZATION_NO;
                worksheet.Cells[i, 11].Value = reports[n].Service_Name;
                worksheet.Cells[i, 12].Value = reports[n].REFERENCE_NO;

                worksheet.Cells[i, 13].Value = reports[n].PAYMENTLOCATION;
                worksheet.Cells[i, 14].Value = reports[n].Transaction_Status;

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


                this.Close();

                excelApp.Quit();
                //if(saveFileDialog1.ShowDialog() )
                System.Diagnostics.Process.Start(saveFileDialog1.FileName);
            }
            else
            {
                this.Close();
            }
          

        }
        public class strs
        {
            public int ACCOUNT_NUMBER { get; set; }
            public string CUSTOMER_NAME { get; set; }

        }

    }
}
