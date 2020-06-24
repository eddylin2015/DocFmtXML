using Newtonsoft.Json;
using System;
using System.Data;
using System.IO;

namespace DocFmtXML
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            // b01c_form.showDocx(  @"C:\code\DocFmtXML\xout0.docx" );
            //   b01c_form.ex();

            String Tml_Doc = @"C:\code\DocFmtXML\DSEJ-B01c_B.docx";
            string strDoc1 = @"C:\code\DocFmtXML\xout.docx";
            String json = "";
            if (File.Exists(@"C:\code\DocFmtXML\td.json")) json = System.IO.File.ReadAllText(@"C:\code\DocFmtXML\td.json");
            DataTable dt = JsonConvert.DeserializeObject<DataTable>(json.Replace("'", "\""));
            DOCF_DSEJB01_FORM.docx(dt, Tml_Doc, strDoc1);
           // DataTable _dt=StafEvalCrossDT.GetCrossTable_type(0);
          //  Console.WriteLine(_dt.Rows.Count);
        }
    }
}
