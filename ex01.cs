using System;
using System.Collections.Generic;
using System.Text;

namespace DocFmtXML
{
    class ex01
    {

        public static void ex()
        {
            //string strDoc = @"c:\temp\DSEJ-B01c.docx";
            string strDoc = @"c:\temp\simple.docx";
            string strDoc1 = @"c:\temp\simple1.docx";
            string txt = "Append text in body - OpenAndAddToWordprocessingStream";
            //Stream stream = File.Open(strDoc, FileMode.Open);
            if (File.Exists(strDoc1)) File.Delete(strDoc1);
            using (Stream outfs = File.Open(strDoc1, FileMode.OpenOrCreate))
            {
                using (FileStream fs = new FileStream(strDoc, FileMode.Open, FileAccess.Read))
                {
                    fs.CopyTo(outfs);
                    fs.Close();
                }
                OpenAndAddToWordprocessingStream_(outfs, txt);
                outfs.Close();
            }
        }
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
        static void ChangeTextInCell(Table table, int rindex, int cindex, String txt)
        {
            TableRow row = table.Elements<TableRow>().ElementAt(rindex);
            TableCell cell = row.Elements<TableCell>().ElementAt(cindex);
            Paragraph p = cell.Elements<Paragraph>().First();
            Run r = p.Elements<Run>().First();
            Text t = r.Elements<Text>().First();
            t.Text = txt;
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
                    Console.WriteLine(run.InnerXml);

                    if (run.InnerText.Contains("FORMCHECKBOX"))
                    {
                        foreach (FieldCode fc in run.Elements<FieldCode>())
                        {
                            Console.Write(fc.InnerXml); Console.Write(" 1* ");
                            foreach (FormFieldData ck in fc.Elements<FormFieldData>())
                            {
                                Console.Write(ck.InnerXml); Console.Write(" 1.1*");
                            }
                        }
                    }
                    else
                    {
                        foreach (FieldChar fc in run.Elements<FieldChar>())
                        {
                            if (fc.FormFieldData != null)
                                Console.Write(fc.FormFieldData.InnerText); Console.Write(" 2*");

                            foreach (FormFieldData ck in fc.Elements<FormFieldData>())
                            {
                                Console.Write(ck.InnerText); Console.Write(" 2.2*");
                            }
                        }
                        foreach (Text fc in run.Elements<Text>())
                        {
                            Console.Write(fc.InnerText); Console.Write(" 4*");

                        }

                        //Console.Write(run.InnerText);
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
