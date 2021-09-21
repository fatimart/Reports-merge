using System;

namespace BatelcoReport
{
    public class Report
    {
        public string ACCOUNT_NUMBER { get; set; }
        public string CUSTOMER_NAME { get; set; }

        public string TRANSACTION_NUMBER { get; set; }

        public DateTime PAYMENTDATE { get; set; }

        public DateTime Date_of_payment_execution { get; set; }
        public double AMOUNT { get; set; }

        public double Commission { get; set; }
        public double VAT { get; set; }
        public double Net_Amount { get; set; }
        public int AUTHRIZATION_NO { get; set; }

        public string Service_Name{ get; set; }

        public string REFERENCE_NO { get; set; }

        public string PAYMENTLOCATION { get; set; }

        public string Transaction_Status { get; set; }

 
    }
}



