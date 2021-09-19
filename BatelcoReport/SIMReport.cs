using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BatelcoReport
{
    public class SIMReport
    {
        public string ACCOUNT_NUMBER { get; set; }
        public string CUSTOMER_NAME { get; set; }

        public string TRANSACTION_NUMBER { get; set; }

        public DateTime PAYMENTDATE { get; set; }

        public DateTime Date_of_payment_execution { get; set; }
        public double PRODUCTAMOUNT { get; set; }

        public double Commission { get; set; }
        public double VATAMOUNT { get; set; }
        public double Net_Amount { get; set; }
        public int KIOSKID { get; set; }

        public string PRODUCTNAME { get; set; }

        public string ORDERNUMBER { get; set; }

        public string PAYMENTLOCATION { get; set; }

        public string Transaction_Status { get; set; }

    }
}
