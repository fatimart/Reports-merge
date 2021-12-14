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
using System.Globalization;

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
        public ObservableCollection<BillReport> TAMReports { get; } = new ObservableCollection<BillReport>();

        public ObservableCollection<MposReport> MposReport { get; } = new ObservableCollection<MposReport>();

        public ObservableCollection<strs> str { get; } = new ObservableCollection<strs>();

        public Merge_Excel_Report ()
        {
            InitializeComponent();

        }

        private void btnSelectFile_Click ( object sender, RoutedEventArgs e )
        {
            OpenKisokFile();

        }

        private void billsBtnBrowse_Click ( object sender, RoutedEventArgs e )
        {
            OpenBillFile();
        }

        private void mPOsBtn_Click ( object sender, RoutedEventArgs e )
        {
            OpenmPOSFile();
        }

        private void btnExport_Click ( object sender, RoutedEventArgs e )
        {
            if (CombineWorkSheet())
            {
                ReadAll();
            }

            


        }



        public void OpenBillFile ()
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Filter = "(.xlsx)|*.xlsx;*.CSV;*.csv";
            fdlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (fdlg.ShowDialog() == true)
            {



                billsFiletxt.Text = fdlg.FileName;

            }
        }


        public void OpenKisokFile ()
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Filter = "(.xlsx)|*.xlsx;*.CSV;*.csv";
            fdlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (fdlg.ShowDialog() == true)
            {

              

                    kioskFiletxt.Text = fdlg.FileName;

                
            }
        }


        public void OpenmPOSFile ()
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Filter = "(.xlsx)|*.xlsx;*.CSV;*.csv";
            fdlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (fdlg.ShowDialog() == true)
            {

               

                    mPOSfileTxt.Text = fdlg.FileName;

                
            }
        }





        public bool CombineWorkSheet ()
        {
            

            string bills, kiosk, mPos;

            bills = billsFiletxt.Text;
            kiosk = kioskFiletxt.Text;
            mPos = mPOSfileTxt.Text;


            if (bills != "" && kiosk != "" && mPos != "")
               {

                    // Create destination Workbook.
                    Aspose.Cells.Workbook destWorkbook = new Aspose.Cells.Workbook();
                    // First worksheet is added by default to the Workbook. Add the second worksheet.
                    destWorkbook.Worksheets.Add();
                    destWorkbook.Worksheets.Add();
                   
                   
                    Aspose.Cells.Workbook SourceBook1 = new Aspose.Cells.Workbook(bills);
                    destWorkbook.Worksheets[0].Copy(SourceBook1.Worksheets[0]);
                    destWorkbook.Worksheets[0].Name = "Sheet1";

                   
                
                    // Open Excel B file.
                    Aspose.Cells.Workbook SourceBook2 = new Aspose.Cells.Workbook(kiosk);
                    destWorkbook.Worksheets[1].Copy(SourceBook2.Worksheets[0]);
                    destWorkbook.Worksheets[1].Name = "Sheet2";

                   
                    // Open the third excel file.
                    Aspose.Cells.Workbook SourceBook3 = new Aspose.Cells.Workbook(mPos);
                    destWorkbook.Worksheets[2].Copy(SourceBook3.Worksheets[0]);
                    destWorkbook.Worksheets[2].Name = "Sheet3";

                    // Save the destination file.
                    destWorkbook.Save("CombinedFile1.xlsx");
                    //System.Diagnostics.Process.Start("CombinedFile1.xlsx");

                    return true;
                }

                else
                {
                   MessageBox.Show("Please select 3 files");
                    return false;
                }

        }

        public bool read1 ()
        {//bills
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
                    if (reader["Transaction Status "].ToString().Trim() == "SUCCESS" && reader["Operation Status "].ToString().Trim() == "COMPLETED" &&  reader["Operation Type "].ToString().Trim() == "BILL_PAYMENT"&& (reader["Service Name "].ToString().Trim() == "BATELCO" || reader["Service Name "].ToString().Trim() == "Batelco"))
                    {
                       
                            string PaymentDate = reader["Transaction Date "].ToString().Trim() + " " + reader["Transaction Time "].ToString().Trim();
               
                        reports.Add(new Report
                            {
                                ACCOUNT_NUMBER = reader["Bill Phone Number "].ToString(),
                                CUSTOMER_NAME = default,
                                TRANSACTION_NUMBER = reader["Reference Number Provider "].ToString(),

                                PAYMENTDATE = Convert.ToDateTime(PaymentDate),
                                Date_of_payment_execution = default,
                                AMOUNT = Convert.ToDouble(reader["Operation Amount "]),

                                Commission = default,
                                VAT = default,
                                Net_Amount = Convert.ToDouble(reader["Operation Amount "]),
                                AUTHRIZATION_NO = default,

                                Service_Name = "TAM",

                                REFERENCE_NO = reader["Transaction Id "].ToString(),

                                PAYMENTLOCATION = "YQB",

                                Transaction_Status = reader["Transaction Status "].ToString().Trim(),
                            });
                        
                    }

                }
                

                connection.Close();
                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show("ERROR in file 2: The file you have selected is not valid. Please check the file path 2 and Select a valid TAM Bills file");
                connection.Close();
                return false;
            }


        }
        public bool read2 ()
        {//kiosk
            string CombineFile = "CombinedFile1.xlsx";
            string commandText = "SELECT * FROM [Sheet2$]";
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

                    
                        if ((reader["BT Res"].ToString().Trim() == "Success" || reader["BT Res"].ToString().Trim() == "Authorization Success") && reader["Service Name"].ToString().Trim() == "Batelco Postpaid")
                        {

                       
                        reports.Add(new Report
                            {
                                ACCOUNT_NUMBER = reader["Phone Number"].ToString().Trim(),
                                CUSTOMER_NAME = default,
                                TRANSACTION_NUMBER = reader["Batelco Transaction ID"].ToString().Trim(),

                                PAYMENTDATE = Convert.ToDateTime(reader["Dateof Payment Received"]),

                                Date_of_payment_execution = default,
                                AMOUNT = Convert.ToDouble(reader["Trx Amount"]),

                                Commission = Convert.ToDouble(reader["Commission"]),
                                VAT = default,
                                Net_Amount = Convert.ToDouble(reader["Net Amount"]),
                                AUTHRIZATION_NO = Convert.ToInt32(reader["Terminal ID"]),

                                Service_Name = reader["Channel Name"].ToString().Trim(),

                                REFERENCE_NO = reader["Transaction Refence No"].ToString().Trim(),

                                PAYMENTLOCATION = "YQB",

                                Transaction_Status = reader["BT Res"].ToString().Trim(),
                            });
                        }
                    
                   

                }

                connection.Close();
                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show("ERROR in file 1: Please Select a valid Kiosk file");
                connection.Close();
                return false;

            }
        }

        public bool read3 ()
        {
            string CombineFile = "CombinedFile1.xlsx";
            string commandText = "SELECT * FROM [Sheet3$]";
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

                    if (reader["Transaction Status"].ToString().Trim() == "Success" && reader["Service Provider"].ToString().Trim() == "Batelco"  && reader["Service Name"].ToString().Trim() == "Batelco Postpaid")
                    {

                        double amount;
                        double.TryParse(reader["YQ Transactio ID"].ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out amount);
                        string stringRepresentationOfDoubleValue = amount.ToString("###");

                        string transactionId = reader["YQ Transactio ID"].ToString();
                        string refNO;
                        if (stringRepresentationOfDoubleValue == "") { refNO = transactionId; }
                        else { refNO = stringRepresentationOfDoubleValue; }
                        // double refg = Convert.ToDouble(reader["YQ Transactio ID"]);
                        // string dateChange = reader["Date of Payment Recieved"].ToString();
                        string outputTime = reader["Date of Payment Recieved"].ToString().Trim();
                        string dateChange = reader["Date of Payment Recieved"].ToString().Trim();
                        string timeFormat = dateChange.Substring(dateChange.Length - 2);
                        if (timeFormat.Equals("AM")|| timeFormat.Equals("am")
                            || timeFormat.Equals("PM")|| timeFormat.Equals("pm"))
                        {
                            switch (timeFormat)
                            {
                                case ("AM"):
                                    outputTime = dateChange.Replace("AM", "");
                                    break;
                                case ("am"):
                                    outputTime = dateChange.Replace("am", "");
                                    break;
                                case ("PM"):
                                    outputTime = dateChange.Replace("PM", "");
                                    break;
                                case ("pm"):
                                    outputTime = dateChange.Replace("pm", "");
                                    break;
                            }
                        }
                        //  String.Format("dd/MM/yyyy HH:mm:ss", dateChange);
                        // DateTime dateEdit = Convert.ToDateTime(dateChange);
                        // string.Format("dd/MM/yyyy HH:mm:ss", dateEdit);

                        // dateEdit.ToString("dd/MM/yyyy HH:mm:ss");
                        //  MessageBox.Show(dateEdit.ToString());

                        DateTime dt = Convert.ToDateTime(outputTime);

                        MposReport.Add(new MposReport
                        {
                            ACCOUNT_NUMBER = reader["Circuit Number"].ToString(),
                            CUSTOMER_NAME = default,
                            TRANSACTION_NUMBER = reader["Batelco Transaction ID"].ToString(),

                           PAYMENTDATE = Convert.ToDateTime(outputTime),
                           //PAYMENTDATE = dt,

                           Date_of_payment_execution = default,
                            AMOUNT = Convert.ToDouble(reader["Trx Amount"]),

                            Commission = Convert.ToDouble(reader["Commission"]),
                            VAT = default,
                            Net_Amount = Convert.ToDouble(reader["Net Amount"]),
                            AUTHRIZATION_NO = default,

                            Service_Name = reader["Channel Name"].ToString(),


                            REFERENCE_NO = refNO,

                            PAYMENTLOCATION = "YQB",

                            Transaction_Status = reader["Transaction Status"].ToString(),
                        });
                    }
                  
                }


                connection.Close();
                return true;

            }

            catch (Exception ex)
            {
                MessageBox.Show("ERROR in file 3: Please Select a valid mPOS file");
                connection.Close();
                return false;

            }
        }
       
        public bool ReadAll ()
        {
            if ( read1() && read2() &&read3())
            {
                Excel();
                return true;
            }
            else
            {
                //MessageBox.Show("ERROR in reading the files: Please Select valid files");
                return false;
            }
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
            int c = reports.Count + MposReport.Count;




            //cells size
            worksheet.Range["A1"].ColumnWidth = 21.43;
            worksheet.Range["B1"].ColumnWidth = 45.14;
            worksheet.Range["C1"].ColumnWidth = 45.43;
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
            worksheet.Cells[1, 1] = "ACCOUNT_NUMBER";
            worksheet.Cells[1, 2] = "CUSTOMER_NAME";
            worksheet.Cells[1, 3] = "TRANSACTION_NUMBER";
            worksheet.Cells[1, 4] = "Transaction Date";
            worksheet.Cells[1, 5] = "Date of payment execution";
            worksheet.Cells[1, 6] = "AMOUNT";
            worksheet.Cells[1, 7] = "Commission";
            worksheet.Cells[1, 8] = "VAT";
            worksheet.Cells[1, 9] = "Net Amount";
            worksheet.Cells[1, 10] = "AUTHRIZATION_NO";
            worksheet.Cells[1, 11] = "Service Name";
            worksheet.Cells[1, 12] = "REFERENCE_NO";
            worksheet.Cells[1, 13] = "Payment Type/Channel Name";
            worksheet.Cells[1, 14] = "Transaction Status";

            worksheet.get_Range("A2", "N" + c + "").NumberFormat = "@";
            worksheet.get_Range("F2", "F" + c + "").NumberFormat = "#,##0.000";
            worksheet.get_Range("G2", "G" + c + "").NumberFormat = "#,##0.000";
            worksheet.get_Range("H2", "H" + c + "").NumberFormat = "#,##0.000";
            worksheet.get_Range("I2", "I" + c + "").NumberFormat = "#,##0.000";


            worksheet.get_Range("A1:N1").Font.Bold = true;
            worksheet.get_Range("A1", "N1").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.get_Range("A1", "N1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            worksheet.get_Range("A2", "N" + c + "").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            worksheet.get_Range("A2", "N" + c + "").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;

            worksheet.get_Range("A1", "N" + c + "").Borders.LineStyle = ex.XlLineStyle.xlContinuous;
            worksheet.get_Range("A1", "N" + c + "").Borders.Weight = ex.XlBorderWeight.xlThin;


            Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;
            Microsoft.Office.Interop.Excel.Range cell = range.Cells[1, 14];
            Microsoft.Office.Interop.Excel.Borders border = cell.Borders;




            for (int k = 0, n = 0; k <= reports.Count - 1; k++, n++)
            {
                int i = k + 2;

                worksheet.Cells[i, 1].Value = "973" + reports[n].ACCOUNT_NUMBER.Trim();
                worksheet.Cells[i, 2].Value = "";
                worksheet.Cells[i, 3].Value = reports[n].TRANSACTION_NUMBER.ToString().Trim();

                worksheet.Cells[i, 4].Value = reports[n].PAYMENTDATE.ToString("dd/MM/yyyy hh:mm:ss").Trim();

                worksheet.Cells[i, 5].Value = "";
                worksheet.Cells[i, 6].Value = reports[n].AMOUNT.ToString().Trim();

                reports[n].Commission = (reports[n].AMOUNT * 0.01);

                decimal rounded_com;
                if (reports[n].Commission >= 0.75)
                { 
                    reports[n].Commission = 0.75;

                    rounded_com = Math.Round(Convert.ToDecimal(reports[n].Commission), 3);
                    worksheet.Cells[i, 7].Value = (rounded_com).ToString().Trim();
                    reports[n].Commission = Convert.ToDouble(rounded_com);
                }

                else
                {
                    rounded_com = Math.Round(Convert.ToDecimal(reports[n].Commission), 3);
                    worksheet.Cells[i, 7].Value = (rounded_com).ToString().Trim();
                    reports[n].Commission = Convert.ToDouble(rounded_com);

                    //worksheet.Cells[i, 7].Value = reports[n].Commission.ToString().Trim(); 
                }


                reports[n].VAT = (0.05 * reports[n].Commission);
                decimal rounded_3 = Math.Round(Convert.ToDecimal(reports[n].VAT), 3);
             
                worksheet.Cells[i, 8].Value = rounded_3.ToString().Trim();

                worksheet.Cells[i, 9].Value = (reports[n].AMOUNT - Convert.ToDouble(rounded_com) - Convert.ToDouble(rounded_3)).ToString().Trim();

                worksheet.Cells[i, 10].Value = "";
                worksheet.Cells[i, 11].Value = reports[n].Service_Name.Trim();
                worksheet.Cells[i, 12].Value = reports[n].REFERENCE_NO.Trim();

                worksheet.Cells[i, 13].Value = reports[n].PAYMENTLOCATION.Trim();
                worksheet.Cells[i, 14].Value = "Success";

            }




            int j = reports.Count  +2;

            for (int i = j, n = 0; n < MposReport.Count - 1; n++, i++)
            {
                worksheet.Cells[i, 1].Value = "973" + MposReport[n].ACCOUNT_NUMBER.Trim();
                worksheet.Cells[i, 2].Value = "";
                worksheet.Cells[i, 3].Value = MposReport[n].TRANSACTION_NUMBER.ToString().Trim();
                worksheet.Cells[i, 4].NumberFormat = "dd/MM/yyyy hh:mm:ss";
                worksheet.Cells[i, 4].Value = MposReport[n].PAYMENTDATE.ToString("dd/MM/yyyy hh:mm:ss").Trim();

                worksheet.Cells[i, 5].Value = "";
                worksheet.Cells[i, 6].Value = MposReport[n].AMOUNT.ToString().Trim();

                MposReport[n].Commission = (MposReport[n].AMOUNT * 0.01);

                decimal rounded_com;
                if (MposReport[n].Commission >= 0.75)
                {
                    MposReport[n].Commission = 0.75;

                    rounded_com = Math.Round(Convert.ToDecimal(MposReport[n].Commission), 3);
                    worksheet.Cells[i, 7].Value = (rounded_com).ToString().Trim();
                    MposReport[n].Commission = Convert.ToDouble(rounded_com);
                }

                else
                {
                    rounded_com = Math.Round(Convert.ToDecimal(MposReport[n].Commission), 3);
                    worksheet.Cells[i, 7].Value = (rounded_com).ToString().Trim();
                    MposReport[n].Commission = Convert.ToDouble(rounded_com);

                    //worksheet.Cells[i, 7].Value = reports[n].Commission.ToString().Trim(); 
                }


                MposReport[n].VAT = (0.05 * MposReport[n].Commission);
                decimal rounded_3 = Math.Round(Convert.ToDecimal(MposReport[n].VAT), 3);

                worksheet.Cells[i, 8].Value = rounded_3.ToString().Trim();

                worksheet.Cells[i, 9].Value = (MposReport[n].AMOUNT - Convert.ToDouble(rounded_com) - Convert.ToDouble(rounded_3)).ToString().Trim();

                worksheet.Cells[i, 10].Value = "";
                worksheet.Cells[i, 11].Value = MposReport[n].Service_Name.Trim();
                worksheet.Cells[i, 12].Value = MposReport[n].REFERENCE_NO.Trim();

                worksheet.Cells[i, 13].Value = MposReport[n].PAYMENTLOCATION.Trim();
                worksheet.Cells[i, 14].Value = "Success";


            }

            //for (int i = j, n = 0; n < BillsReports.Count - 1; n++, i++)
            //{
            //    worksheet.Cells[i, 1].Value = "973" + BillsReports[n].ACCOUNT_NUMBER.Trim();
            //    worksheet.Cells[i, 2].Value = BillsReports[n].CUSTOMER_NAME.Trim();
            //    worksheet.Cells[i, 3].Value = BillsReports[n].TRANSACTION_NUMBER.ToString().Trim();

            //    worksheet.Cells[i, 4].Value = BillsReports[n].PAYMENTDATE.ToString().Trim();
            //    //worksheet.Cells[i, 4].Value = BillsReports[n].PAYMENTDATE.ToString("dd/MM/yyyy HH:mm:ss").Trim();

            //    worksheet.Cells[i, 5].Value = "";
            //    worksheet.Cells[i, 6].Value = BillsReports[n].PRODUCTAMOUNT.ToString().Trim();

            //    BillsReports[n].Commission = (BillsReports[n].PRODUCTAMOUNT * 0.01);

            //    if (BillsReports[n].Commission >= 0.750)
            //    {
            //        BillsReports[n].Commission = 0.75;
            //        worksheet.Cells[i, 7].Value = BillsReports[n].Commission.ToString().Trim();
            //    }
            //    else
            //    {
            //        worksheet.Cells[i, 7].Value = BillsReports[n].Commission.ToString().Trim();
            //    }


            //    worksheet.Cells[i, 8].Value = BillsReports[n].VATAMOUNT.ToString().Trim();


            //    worksheet.Cells[i, 9].Value = (BillsReports[n].PRODUCTAMOUNT - BillsReports[n].Commission).ToString().Trim();
            //    worksheet.Cells[i, 10].Value = BillsReports[n].KIOSKID.ToString().Trim();
            //    worksheet.Cells[i, 11].Value = BillsReports[n].PRODUCTNAME.Trim();
            //    worksheet.Cells[i, 12].Value = BillsReports[n].ORDERNUMBER.Trim();
            //    worksheet.Cells[i, 13].Value = BillsReports[n].PAYMENTLOCATION.Trim();
            //    worksheet.Cells[i, 14].Value = "Success";


            //}

            //int C1 = j + BillsReports.Count -2 ;
            //worksheet.get_Range("A" + (j - 1) + "", "N" + (j - 1) + "").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //worksheet.get_Range("A" + (j - 1) + "", "N" + (j - 1) + "").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            //worksheet.get_Range("A" + j + "", "N" + C1 + "").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //worksheet.get_Range("A" + j + "", "N" + C1 + "").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;

            //worksheet.get_Range("A" + (j - 1) + "", "N" + C1 + "").Borders.LineStyle = ex.XlLineStyle.xlContinuous;
            //worksheet.get_Range("A" + (j - 1) + "", "N" + C1 + "").Borders.Weight = ex.XlBorderWeight.xlThin;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Save Excel sheet";
            saveFileDialog1.Filter = "Excel files|*.xlsx|All files|*.*";
            saveFileDialog1.FileName = "YQ_Sample_Report " + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx";


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
                //System.Diagnostics.Process.Start(saveFileDialog1.FileName);
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
