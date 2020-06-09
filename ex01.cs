using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace DocFmtXML
{
    class ex01
    {
        static string json = @"
[{
        'STUD_ID': '7E39999A',
        'CODE': '1234567-X',
        'YEAR': '2019/2020',
        'BP': '2',
        'IP': '1',
        'IDT': 'BIRNP',
        'RAR': 'M',
        'AR': 'M',
        'GAR': 'M',
        'EAR': 'M',
        'ZH_S_CODE': '澳門中學',
        'PRE_S_CODE': '',
        'NAME_C': '周星星',
        'NAME_P': 'SOU XING XING',
        'SEX': 'M',
        'B_DATE': '2012/01/01',
        'B_PLACE': '',
        'ID_TYPE': '',
        'ID_NO': '1006010(0)',
        'I_PLACE': '',
        'I_DATE': '2019/01/07',
        'V_DATE': '2024/01/07',
        'S6_TYPE': '3=其他逗留許可',
        'S6_IDATE': null,
        'S6_VDATE': null,
        'NATION': '中國',
        'ORIGIN': '廣東',
        'R_AREA': '',
        'RA_DESC': '',
        'AREA': 'M=澳門',
        'POSTAL_CODE': '',
        'ROAD': '大馬路',
        'ADDRESS': '新邨N樓X座',
        'TEL': '/',
        'MOBILE': '61111177',
        'FATHER': '周大福',
        'MOTHER': '秀梅',
        'F_PROF': '律師',
        'M_PROF': '主婦',
        'GUARD': 'M',
        'LIVE_SAME': '0',
        'EC_NAME': '秀梅',
        'EC_REL': '母子',
        'EC_TEL': '61111137',
        'EC_AREA': 'M=澳門',
        'EC_POSTAL_CODE': '',
        'EC_ROAD': '大馬路',
        'EC_ADDRESS': '新邨N樓X座',
        'S_CODE': '159',
        'GRADE': 'P1',
        'CLASS': 'A',
        'C_NO': '65',
        'G_NAME': '秀梅',
        'G_RELATION': '',
        'G_PROFESSION': '主婦',
        'G_AREA': 'M=澳門',
        'G_POSTAL_CODE': '',
        'G_ROAD': '大馬路',
        'G_ADDRESS': '新邨N樓X座',
        'G_TEL': '61111177',
        'GUARDMOBIL': '61111177',
        'F_tel1': '61111197',
        'F_tel2': '/',
        'M_tel1': '61111177',
        'M_tel2': '/',
        'G_tel1': '61111377',
        'G_tel2': '/',
        'Parent_sms': '61111377',
        'Stud_sms': '',
        'Reg_in_date': '2019-09-09',
        'Reg_in_Class': 'P1A',
        'St_status': '3=插班',
        'Leave_date': '',
        'Leave_Class': '',
        'Leave_reason': '',
        'Religion': '',
        'MBC_STUD': '0',
        'K_CLASS': '',
        'K_SCHOOL': '',
        'K_EDU': '',
        'P_CLASS': 'P2',
        'P_SCHOOL': '台東小學',
        'P_EDU': '',
        'S_CLASS': '',
        'S_SCHOOL': '',
        'S_EDU': '',
        'note': '',
        'last_class': ''
    }]
";
        public static void ex()
        {
            String Tml_Doc = @"C:\code\DocFmtXML\DSEJ-B01c_N.docx";
            string strDoc1 = @"C:\code\DocFmtXML\simple1.docx";
            //Stream stream = File.Open(strDoc, FileMode.Open);
            string json = System.IO.File.ReadAllText(@"c:\temp\td.json");
            if (File.Exists(@"c:\temp\td.json"))
                json = System.IO.File.ReadAllText(@"c:\temp\td.json");
            DataTable dt = JsonConvert.DeserializeObject<DataTable>(json.Replace("'", "\""));
            if (File.Exists(strDoc1)) File.Delete(strDoc1);
            using (Stream outfs = File.Open(strDoc1, FileMode.OpenOrCreate))
            {
                using (FileStream fs = new FileStream(Tml_Doc, FileMode.Open, FileAccess.Read))
                {
                    fs.CopyTo(outfs);
                    fs.Close();
                }
                WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(outfs, true);
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                List<OpenXmlElement> templete_li = new List<OpenXmlElement>();
                foreach (var ele in body.ChildElements)
                {
                    templete_li.Add(ele);
                }
                foreach (var ele in templete_li)
                {
                    Console.WriteLine(ele.ToString());
                    if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.Table"))
                    {
                        DocumentFormat.OpenXml.Wordprocessing.Table _tbl = (DocumentFormat.OpenXml.Wordprocessing.Table)ele;
                        showTable(_tbl);
                        //Console.WriteLine(_tbl.InnerText);
                        // if (_tbl.InnerText.Contains("上學年度")) { ChangeTextInCell(_tbl, 0, 1, pagecnt.ToString()); }
                        if (_tbl.InnerText.Contains("學生個人資料"))
                        {
                            //ChangeChkBox(_tbl, 2, 1, 1);
                        }
                    }
                    if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.Paragraph"))
                    {
                        //Paragraph _prg = (Paragraph)ele;
                        //Console.WriteLine(_prg.InnerText);
                    }
                    if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.SectionProperties")) { }
                }
                
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    Paragraph para = body.AppendChild(new Paragraph(new Run((new Break() { Type = BreakValues.Page }))));
                    List<OpenXmlElement> clone_li = new List<OpenXmlElement>();

                    foreach (var ele in templete_li)
                    {
                        clone_li.Add((OpenXmlElement)ele.Clone());
                    }
                    fillRow(clone_li, dt.Rows[i]);
                    body.Append(clone_li);
                }
                fillRow(templete_li, dt.Rows[0]);

                wordprocessingDocument.Close();
                outfs.Close();
            }
        }
        static void fillRow(List<OpenXmlElement> li,DataRow dr)
        {
            foreach (var ele in li)
            {
                if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.Table"))
                {
                    if (ele.InnerText.Contains("首次註冊"))
                    {
                        var table_ = (Table)ele;
                        string[] year_arr = dr["YEAR"].ToString().Split('/');
                        ChangeTextInCell(table_, 0, 1, year_arr[0]);
                        ChangeTextInCell(table_, 0, 3, year_arr[1]);
                        char[] code_arr = dr["CODE"].ToString().ToCharArray();
                        ChangeTextInCell(table_, 0, 5, code_arr[0].ToString());
                        ChangeTextInCell(table_, 0, 7, code_arr[1].ToString());
                        ChangeTextInCell(table_, 0, 9, code_arr[2].ToString());
                        ChangeTextInCell(table_, 0, 11, code_arr[3].ToString());
                        ChangeTextInCell(table_, 0, 13, code_arr[4].ToString());
                        ChangeTextInCell(table_, 0, 15, code_arr[5].ToString());
                        ChangeTextInCell(table_, 0, 17, code_arr[6].ToString());
                        ChangeTextInCell(table_, 0, 19, code_arr[8].ToString());
                    }
                    else if (ele.InnerText.Contains("上學年度"))
                    {
                        if (dr["last_class"].ToString().Length > 2){
                            var table_ = (Table)ele;
                            ChangeTextInCell(table_, 0, 1, String.Format("159  澳門浸信中學      ({0})", dr["last_class"].ToString()));
                        }
                    }
                    
                else if (ele.InnerText.Contains("註冊資料"))
                    {
                        var table_ = (Table)ele;
                        ChangeTextInCell(table_, 0, 2, "159");
                        ChangeTextInCell(table_, 0, 4, "澳門浸信中學");
                        ChangeTextInCell(table_, 1, 2, dr["GRADE"].ToString());
                        ChangeTextInCell(table_, 1, 4, dr["CLASS"].ToString());
                        ChangeTextInCell(table_, 1, 6, dr["C_NO"].ToString());
                    }
                    else if (ele.InnerText.Contains("學生個人資料"))
                    {
                        var table_ = (Table)ele;
                        ChangeTextInCell(table_, 0, 2, dr["NAME_C"].ToString());
                        ChangeTextInCell(table_, 0, 4, dr["NAME_P"].ToString());
                        ChangeDateInCell(table_, 1, 3, dr["B_DATE"].ToString());
                        if (dr["SEX"].ToString().Equals("M"))
                        {
                            SetChkBox(table_, 2, 1, 0);
                        }
                        else if (dr["SEX"].ToString().Equals("F"))
                        {
                            SetChkBox(table_, 2, 1, 1);
                        }
                    } 
                }
            }
        }
        static void SetChkBox(Table table, int rindex, int cindex, int i)
        {
            int cnt = 0;
            TableCell cell = GetCell(table, rindex, cindex);
            foreach (Paragraph parag in cell.Elements<Paragraph>())
            {
                foreach (Run run in parag.Elements<Run>())
                {
                    if (!run.InnerXml.Contains("check")) continue;
                    //Console.WriteLine(run.InnerXml);
                    //foreach (FieldCode fc in run.Elements<FieldCode>())  Console.Write(fc.InnerXml); Console.Write(" 1* ");
                    foreach (FieldChar fc in run.Elements<FieldChar>())
                        {
                        if (fc.FormFieldData != null)
                        {
                            if (cnt == i)
                            {
                                run.InnerXml = run.InnerXml.Replace("w:val=\"0\"", "w:val=\"1\"");
                            }
                            cnt++;
                            
                        }
                        foreach (FormFieldData ck in fc.Elements<FormFieldData>())
                        {
                                //Console.Write(ck.InnerText); Console.Write(" 2.2*");
                        }
                    }
                }
            }
        }
        public static void out_B01c(Stream stream, DataTable dt)
        {
            // Open a WordProcessingDocument based on a stream.
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true);
            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            // Add new text.
            //Paragraph para = body.AppendChild(new Paragraph());
            //Run run = para.AppendChild(new Run());
            //run.AppendChild(new Text(txt));
            //my coding 
            List<OpenXmlElement> templete_li = new List<OpenXmlElement>();
            foreach (var ele in body.ChildElements)
            {
                templete_li.Add((OpenXmlElement)ele.Clone());
            }

            if(dt.Rows[0]["St_status"].ToString().Equals("1=新生"))
            {

            }
            
            for (int i = 1; i < 2; i++)
            {
                Paragraph para = body.AppendChild(new Paragraph(new Run((new Break() { Type = BreakValues.Page }))));
                List<OpenXmlElement> clone_li = new List<OpenXmlElement>();
                foreach (var ele in templete_li)
                {
                    clone_li.Add((OpenXmlElement)ele.Clone());
                }
                body.Append(clone_li);
            }
            int pagecnt = 0;
            foreach (var ele in body.ChildElements)
            {
                Console.WriteLine(ele.ToString());
                if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.Table"))
                {
                    DocumentFormat.OpenXml.Wordprocessing.Table _tbl = (DocumentFormat.OpenXml.Wordprocessing.Table)ele;
                    showTable(_tbl);
                    //Console.WriteLine(_tbl.InnerText);
                   // if (_tbl.InnerText.Contains("上學年度")) { ChangeTextInCell(_tbl, 0, 1, pagecnt.ToString()); }
                    if (_tbl.InnerText.Contains("學生個人資料"))
                    {
                        //ChangeChkBox(_tbl, 2, 1, 1);
                    }
                }
                if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.Paragraph"))
                {
                    Paragraph _prg = (Paragraph)ele;
                    Console.WriteLine(_prg.InnerText);
                }
                if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.SectionProperties")) pagecnt++;
            }

            // Close the document handle.
            wordprocessingDocument.Close();
            // Caller must close the stream.

        }
        /*
         * System.IO.File.WriteAllText(@"C:\code\ds_" + cno + ".json", Newtonsoft.Json.JsonConvert.SerializeObject(ds));
         *
         *string json = System.IO.File.ReadAllText(@"json/ds_" + pclass + ".json");
            DataSet ds = JsonConvert.DeserializeObject<DataSet>(json);
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                Console.WriteLine(ds.Tables[i].TableName);
            }
            DataColumn pcol = ds.Tables["Table"].Columns["stud_ref"];
            String[] subtbNs = { "py", "cd", "mk", "ac", "gc" };
            for (int i = 0; i < subtbNs.Length; i++)
            {
                DataColumn ccol = ds.Tables[subtbNs[i]].Columns["stud_ref"];
                if (ccol != null)
                {
                    DataRelation dr = new DataRelation("sr_" + subtbNs[i], pcol, ccol);
                    dr.Nested = true;
                    ds.Relations.Add(dr);
                }
            }
      */
        public static void ex_open_append_text()
        {
            //string strDoc = @"c:\temp\DSEJ-B01c.docx";
            string strDoc = @"c:\temp\simple.docx";
            string txt = "Append text in body - OpenAndAddToWordprocessingStream";
            Stream stream = File.Open(strDoc, FileMode.Open);
            OpenAndAddToWordprocessingStream(stream, txt);
            stream.Close();
        }
        public static void OpenAndAddToWordprocessingStream_(Stream stream, string txt)
        {
            // Open a WordProcessingDocument based on a stream.
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true);
            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            // Add new text.
            //Paragraph para = body.AppendChild(new Paragraph());
            //Run run = para.AppendChild(new Run());
            //run.AppendChild(new Text(txt));

            //my coding 
            List<OpenXmlElement> templete_li = new List<OpenXmlElement>();
            foreach (var ele in body.ChildElements)
            {
                templete_li.Add(ele);
            }
            for (int i = 1; i < 2; i++)
            {
                Paragraph para = body.AppendChild(new Paragraph(new Run((new Break() { Type = BreakValues.Page }))));
                List<OpenXmlElement> clone_li = new List<OpenXmlElement>();
                foreach (var ele in templete_li)
                {
                    clone_li.Add((OpenXmlElement)ele.Clone());
                }
                body.Append(clone_li);
            }
            int pagecnt = 0;
            foreach (var ele in body.ChildElements)
            {
                Console.WriteLine(ele.ToString());
                if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.Table"))
                {
                    DocumentFormat.OpenXml.Wordprocessing.Table _tbl = (DocumentFormat.OpenXml.Wordprocessing.Table)ele;
                    showTable(_tbl);
                    //Console.WriteLine(_tbl.InnerText);
                    if (_tbl.InnerText.Contains("上學年度")) { ChangeTextInCell(_tbl, 0, 1, pagecnt.ToString()); }
                    if (_tbl.InnerText.Contains("學生個人資料"))
                    {
                        ChangeChkBox(_tbl, 2, 1, 1);
                    }
                }
                if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.Paragraph"))
                {
                    Paragraph _prg = (Paragraph)ele;
                    Console.WriteLine(_prg.InnerText);
                }
                if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.SectionProperties")) pagecnt++;
            }

            // Close the document handle.
            wordprocessingDocument.Close();
            // Caller must close the stream.
        }
        static void WriteCell(Table table, int rindex, int cindex, String txt)
        {
            TableRow row = table.Elements<TableRow>().ElementAt(rindex);
            TableCell cell = row.Elements<TableCell>().ElementAt(cindex);
            Paragraph p = cell.Elements<Paragraph>().First();
            Run r = p.Elements<Run>().First();
            Text t = r.Elements<Text>().First();
            t.Text = txt;
        }
        static void ChangeDateInCell(Table table, int rindex, int cindex, String txt)
        {
            TableRow row = table.Elements<TableRow>().ElementAt(rindex);
            TableCell cell = row.Elements<TableCell>().ElementAt(cindex);
            foreach (OpenXmlElement parag in cell.ChildElements)
            {
                if (parag.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.Table"))
                {
                    Table tbl_ = (Table)parag;
                    String[] arr=txt.Split('/');
                    ChangeTextInCell(tbl_, 0, 0, arr.Length > 0 ? arr[0] :"");
                    ChangeTextInCell(tbl_, 0, 2, arr.Length > 1 ? arr[1] : "");
                    ChangeTextInCell(tbl_, 0, 4, arr.Length > 2 ? arr[2] : "");
                }
            }
        }

        static void ChangeTextInCell(Table table, int rindex, int cindex, String txt)
        {
            TableRow row = table.Elements<TableRow>().ElementAt(rindex);
            TableCell cell = row.Elements<TableCell>().ElementAt(cindex);
            Paragraph p = cell.Elements<Paragraph>().First();
            var r_li = p.Elements<Run>().ToArray();
            if (r_li.Length > 0)
            {
                //Run r = p.Elements<Run>().First();
                Run r = r_li[0];
                Text t = r.Elements<Text>().First();
                t.Text = txt;
            }
            else
            {
                Run run = p.AppendChild(new Run());
                run.AppendChild(new Text(txt));
            }
        }
        static TableCell GetCell(Table table, int rindex, int cindex)
        {
            TableRow row = table.Elements<TableRow>().ElementAt(rindex);
            TableCell cell = row.Elements<TableCell>().ElementAt(cindex);
            return cell;
        }

        static void ChangeChkBox(Table table, int rindex, int cindex, int i)
        {
            TableCell cell = GetCell(table, rindex, cindex);
            //Console.WriteLine(cell.InnerText);
            foreach (Paragraph parag in cell.Elements<Paragraph>())
            {
                foreach (Run run in parag.Elements<Run>())
                {
                    run.InnerXml = run.InnerXml.Replace("<w:checked w:val=\"0\" />", "<w:checked w:val=\"1\" />");
                    //Console.WriteLine(run.InnerXml);
                    if (run.InnerText.Contains("FORMCHECKBOX"))
                    {
                        foreach (FieldCode fc in run.Elements<FieldCode>())
                        {
                            //Console.Write(fc.InnerXml); Console.Write(" 1* ");
                            foreach (FormFieldData ck in fc.Elements<FormFieldData>())
                            {
                                //Console.Write(ck.InnerXml); Console.Write(" 1.1*");
                            }
                        }
                    }
                    else
                    {
                        foreach (FieldChar fc in run.Elements<FieldChar>())
                        {
                            if (fc.FormFieldData != null)
                            {
                             //Console.Write(fc.FormFieldData.InnerText); Console.Write(" 2*");
                            }

                            foreach (FormFieldData ck in fc.Elements<FormFieldData>())
                            {
                             //Console.Write(ck.InnerText); Console.Write(" 2.2*");
                            }
                        }
                        foreach (Text fc in run.Elements<Text>())
                        {
                            //Console.Write(fc.InnerText); Console.Write(" 4*");
                        }
                    }
                }
            }
            /*
           foreach (Paragraph parag in cell.Elements<Paragraph>())
            {
                if (parag.InnerText.Contains("FORMCHECKBOX"))
                {
                    //Console.Write("XV");
                    foreach (Run run in parag.Elements<Run>())
                    {

                        if (run.InnerText.Contains("FORMCHECKBOX"))
                        {
                            Console.Write("XV");
                        }
                        else
                        {
                            Console.Write(run.InnerText);
                        }
                    }
                }
             }*/
    }
    static void showTable(Table _tbl)
        {
            foreach (TableRow row in _tbl.Elements<TableRow>())
            {
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    foreach (Paragraph parag in cell.Elements<Paragraph>())
                    {

                        if (parag.InnerText.Contains("FORMCHECKBOX"))
                        {
                            //Console.Write("XV");
                            foreach (Run run in parag.Elements<Run>())
                            {

                                if (run.InnerText.Contains("FORMCHECKBOX"))
                                {
                                    Console.Write("XV");
                                }
                                else
                                {
                                    Console.Write(run.InnerText);
                                }


                            }
                        }
                        else
                        {
                            Console.Write(parag.InnerText);
                        }




                        //Console.Write("\t");
                        //IEnumerable<Run> runs = parag.Elements<Run>();

                        //Run run=runs.GetEnumerator().Current;
                        //Text text = run.Elements<Text>().GetEnumerator().Current;

                        //Run run = parag.Elements<Run>().First();
                        // Set the text for the run.  
                        //Text text = run.Elements<Text>().First();
                        //text.Text = addedText;

                    }
                    Console.Write("\t");

                }
                Console.WriteLine();
            }
        }
        public static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
        {
            // Open a WordProcessingDocument based on a stream.
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(stream, true);
            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));

            // Close the document handle.
            wordprocessingDocument.Close();
            // Caller must close the stream.
        }
    }
}
