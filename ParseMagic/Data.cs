using System.Collections.Generic;
using System.Data;

namespace ParseMagic
{
    class Data
    {
        public static string _type = "";
        public static List<string> selectors = new List<string>() { "[ ]", "{ }", "< >" };
        public static List<int> sleepTimes = new List<int>() { 50, 100, 200, 300, 500, 1000, 2000 };
        public static string _xcelfilelocation = "";
        public static string _textfilelocation = "";
        public static int _reciepientsCount = 0;
        public static int _attributeCount = 0;
        public static string _baseTextStructure = @"TEL: [TELEPHONE]
CORRECTION: [CUSTOMER NAME]: We have been adviced by the bank that the product name [PRODUCT] captured in our demand letter/sms/email to you on your outstanding indeptedness to the bank on account No. [ACCOUNT NO.] is wrong. The correct product name is [STATUS] and not [PRODUCT].
This mistake is highly regretted. (Text from the Law Firm of P. C. Ebunilo & Co. (Solicitors to Access/Diamond Bank).";

        public static DataSet ds;
        public static DataTable dt;
        public static object[][] dt_as_array;
        public static List<string> dt_as_array_columns = new List<string>();
        public static int rowCount;
        public static int colCount;
        public static List<List<string>> tuples = new List<List<string>>();

        public static List<string> fileAttributesFound = new List<string>();

        public static int _textAttributeCount = 0;
        public static char leftSelector = '[';
        public static char rightSelector = ']';
        public static int processed = 0;

        public static bool _open_after_export = true;

        public static void flush()
        {
            _xcelfilelocation = "";
            _textfilelocation = "";
            _reciepientsCount = 0;
            _attributeCount = 0;

            ds = null;
            dt = null;
            dt_as_array = null;
            dt_as_array_columns.Clear();
            rowCount = 0;
            colCount = 0;
            _textAttributeCount = 0;
            processed = 0;
            fileAttributesFound.Clear();
            tuples.Clear();
        }
    }
}
