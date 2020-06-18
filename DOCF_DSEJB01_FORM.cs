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
    class DOCF_DSEJB01_FORM
    {
        public static void docx(DataTable dt,String Tml_Doc,String strDoc1)
        {
            //String Tml_Doc = @"C:\code\DocFmtXML\DSEJ-B01c_B.docx";
            //string strDoc1 = @"C:\code\DocFmtXML\xout.docx";
            //if (File.Exists(@"C:\code\DocFmtXML\td.json")) json = System.IO.File.ReadAllText(@"C:\code\DocFmtXML\td.json");
            //DataTable dt = JsonConvert.DeserializeObject<DataTable>(json.Replace("'", "\""));
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
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    Paragraph para = body.AppendChild(new Paragraph(new Run((new Break() { Type = BreakValues.Page }))));
                    List<OpenXmlElement> clone_li = new List<OpenXmlElement>();
                    int eleindex = 0;
                    bool newreg_flag = dt.Rows[i]["St_status"].Equals("1=新生");
                    foreach (var ele in templete_li)
                    {
                        eleindex++;
                        if (newreg_flag && eleindex == 34) continue;
                        if (!newreg_flag && eleindex == 32) continue;
                        var clone_ele = (OpenXmlElement)ele.Clone();
                        if (eleindex == 34 || eleindex == 32)
                        {
                            Table reginfo_table = GetCell((Table)clone_ele, 0, 0).Elements<Table>().ElementAt(0);
                            fillTextInTable(reginfo_table, reqinfo_field_posi, dt.Rows[i]);
                        }
                        clone_li.Add(clone_ele);
                    }
                    fillRow(clone_li, dt.Rows[i]);
                    body.Append(clone_li);
                }
                fillRow(templete_li, dt.Rows[0]);
                {
                    Table reginfo_table = GetCell((Table)templete_li.ElementAt(33), 0, 0).Elements<Table>().ElementAt(0);
                    fillTextInTable(reginfo_table, reqinfo_field_posi, dt.Rows[0]);
                }
                {
                    Table reginfo_table = GetCell((Table)templete_li.ElementAt(31), 0, 0).Elements<Table>().ElementAt(0);
                    fillTextInTable(reginfo_table, reqinfo_field_posi, dt.Rows[0]);
                }
                if (dt.Rows[0]["St_status"].Equals("1=新生"))
                {
                    body.RemoveChild(templete_li.ElementAt(33));
                }
                else
                {
                    body.RemoveChild(templete_li.ElementAt(31));
                }
                wordprocessingDocument.Close();
                outfs.Close();
            }
        }
        static void fillTextInTable(Table table_, String[] baseinfo_field_posi, DataRow dr)
        {
            for (int i = 0; i < baseinfo_field_posi.Length / 2; i++)
            {
                if (baseinfo_field_posi[i * 2].ToUpper().Contains("DATE"))
                {
                    string[] arr = baseinfo_field_posi[i * 2 + 1].Split('.');
                    if (arr.Length == 2)
                        ChangeDateInCell(table_, int.Parse(arr[0]), int.Parse(arr[1]), dr[baseinfo_field_posi[i * 2]].ToString());

                }
                else if (baseinfo_field_posi[i * 2].ToUpper().Contains("SEX"))
                {
                }
                else
                {
                    string[] arr = baseinfo_field_posi[i * 2 + 1].Split('.');
                   // Console.WriteLine(baseinfo_field_posi[i * 2]);
                    if (arr.Length == 2)
                        ChangeTextInCell(table_, int.Parse(arr[0]), int.Parse(arr[1]), dr[baseinfo_field_posi[i * 2]].ToString());
                }
            }
        }

        static void fillRow(List<OpenXmlElement> li, DataRow dr)
        {
            int cnt = 0;
            foreach (var ele in li)
            {
                if (ele.ToString().Equals("DocumentFormat.OpenXml.Wordprocessing.Table"))
                {
                    cnt++;
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
                        var table_ = (Table)ele;
                        ChangeTextInCell(table_, 0, 1, String.Format("{0}", dr["PRE_S_CODE"].ToString()));
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
                        fillTextInTable(table_, baseinfo_field_posi, dr);
                        ChangeDateInCell(table_, 1, 3, dr["B_DATE"].ToString());
                        if (dr["SEX"].ToString().Equals("M")) { SetChkBox(table_, 2, 1, 0); }
                        else if (dr["SEX"].ToString().Equals("F")) { SetChkBox(table_, 2, 1, 1); }
                        if (dr["BP"].ToString().Equals("1")) { SetChkBox(table_, 2, 4, 0); }
                        else if (dr["BP"].ToString().Equals("2")) { SetChkBox(table_, 2, 4, 1); }
                        if (dr["IDT"].ToString().Equals("BIRP")) { SetChkBox(table_, 4, 2, 0); }
                        else if (dr["IDT"].ToString().Equals("BIRNP")) { SetChkBox(table_, 4, 2, 1); }
                        if (dr["IP"].ToString().Equals("1")) { SetChkBox(table_, 6, 2, 0); }
                        else if (dr["IP"].ToString().Equals("2")) { SetChkBox(table_, 6, 2, 1); }
                        if (dr["RAR"].ToString().Equals("M")) { SetChkBox(table_, 9, 3, 0); }
                        else if (dr["RAR"].ToString().Equals("C")) { SetChkBox(table_, 9, 3, 1); }
                        if (dr["AR"].ToString().Equals("M")) { SetChkBox(table_, 11, 2, 0); }
                        else if (dr["AR"].ToString().Equals("T")) { SetChkBox(table_, 11, 2, 1); }
                        else if (dr["AR"].ToString().Equals("C")) { SetChkBox(table_, 11, 2, 2); }
                        else if (dr["AR"].ToString().Equals("L")) { SetChkBox(table_, 11, 2, 3); }
                        if (dr["GUARD"].ToString().Equals("F")) { SetChkBox(table_, 16, 3, 0); }
                        else if (dr["GUARD"].ToString().Equals("M")) { SetChkBox(table_, 16, 3, 1); }
                        else if (dr["GUARD"].ToString().Equals("O")) { SetChkBox(table_, 16, 3, 2); }

                    }
                    else if (cnt == 6)
                    {
                        var table_ = (Table)ele;
                        fillTextInTable(table_, GU_field_posi, dr);
                        if (dr["GAR"].ToString().Equals("M")) { SetChkBox(table_, 2, 2, 0); }
                        else if (dr["GAR"].ToString().Equals("C")) { SetChkBox(table_, 2, 2, 1); }

                    }
                    else if (cnt == 7)
                    {
                        var table_ = (Table)ele;
                        fillTextInTable(table_, EC_field_posi, dr);
                        if (dr["EAR"].ToString().Equals("M")) { SetChkBox(table_, 2, 2, 0); }
                        else if (dr["EAR"].ToString().Equals("C")) { SetChkBox(table_, 2, 2, 1); }

                    }
                }
            }
        }
        static TableCell GetCell(Table table, int rindex, int cindex)
        {
            TableRow row = table.Elements<TableRow>().ElementAt(rindex);
            TableCell cell = row.Elements<TableCell>().ElementAt(cindex);
            return cell;
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
                    String[] arr = txt.Split('/');
                    ChangeTextInCell(tbl_, 0, 0, arr.Length > 0 ? arr[0] : "");
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
        static void SetChkBox(Table table, int rindex, int cindex, int i)
        {
            int cnt = 0;
            TableCell cell = GetCell(table, rindex, cindex);
            foreach (Paragraph parag in cell.Elements<Paragraph>())
            {
                foreach (Run run in parag.Elements<Run>())
                {
                    if (!run.InnerXml.Contains("check")) continue;
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
                    }
                }
            }
        }
        public static string[] baseinfo_field_posi = new string[]{
        "NAME_C","0.2",
        "NAME_P","0.4",
        "SEX","2.1",
        "B_DATE","1.3",
        "B_PLACE","1.5",
        "ID_TYPE","3.3",
        "ID_NO","3.5",
        "I_PLACE","5.3",
        "I_DATE","6.4",
        "V_DATE","6.5",
        "S6_TYPE","",
        "S6_IDATE", "7.3",
        "S6_VDATE", "7.5",
        "NATION","9.1",
        "ORIGIN","9.2",
        "R_AREA","",
        "RA_DESC","8.4",
        "AREA","",
        "POSTAL_CODE","10.3",
        "ROAD","10.5",
        "ADDRESS","12.3",
        "TEL","12.5",
        "MOBILE","13.5",
        "FATHER","14.2",
        "MOTHER","15.2",
        "F_PROF","14.4",
        "M_PROF","15.4",
        "GUARD","",
        "G_RELATION","16.4",
        "GUARDMOBIL","16.6",
        "LIVE_SAME",""
       };

        public static string[] GU_field_posi = new string[]{
        "G_NAME","0.2",
        "G_RELATION","",
        "G_PROFESSION","0.4",
        "G_AREA","",
        "G_POSTAL_CODE","",
        "G_ROAD","1.4",
        "G_ADDRESS","3.3",
        "G_TEL","3.5",
        };

        public static string[] EC_field_posi = new string[]{
        "EC_NAME","0.2",
        "EC_REL","0.4",
        "EC_TEL","3.6",
        "EC_AREA","",
        "EC_POSTAL_CODE"," ",
        "EC_ROAD","1.4",
        "EC_ADDRESS","3.3" };
        public static string[] reqinfo_field_posi = new string[]{
        "F_tel1"," 0.1",
        "F_tel2","0.3",
        "M_tel1","1.1",
        "M_tel2","1.3",
        "G_tel1","2.1",
        "G_tel2","2.3",
        "Parent_sms","3.1",
        "STUD_ID","4.1",
        "Stud_sms","5.1" };

    }
}
