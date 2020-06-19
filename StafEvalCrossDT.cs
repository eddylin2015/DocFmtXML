using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace DocFmtXML
{
    public partial class     ESData
    {
        
        
        static ESData _instance = null;
        public static ESData GetInst
        {
            get
            {
                if (_instance == null) { _instance = new ESData(); conn.Open(); }
                return _instance;
            }
        }
        public MySqlDataReader Reader(string sql)
        {
            return new MySqlCommand(sql, conn).ExecuteReader();
        }
    }
    public class StafEvalCrossDT
    {
        //http://localhost:8082/api/StudInfo/echo
        private void ShowCrossValXRate(Object Sender, EventArgs e)
        {
            ShowCrossTable_type(2);
        }

        private void ShowCrossVal(Object Sender, EventArgs e)
        {
            ShowCrossTable_type(1);
        }
        private void ShowCrossAns(Object Sender, EventArgs e)
        {
            ShowCrossTable_type(0);
        }
        private void ShowCrossTable_type(int int_type)
        {
            //ES_FORMS.iDTFromGrid idt = new ES_FORMS.iDTFromGrid(GetCrossTable_type(int_type));
            //ES_FORMS.FormDataGrid_Cmd fdg = new FormDataGrid_Cmd(idt, null, BindingListOptions.AllowModifyNo);
            //fdg.MdiParent = this;
            //fdg.Show();
        }
        private void ShowCrossAnsforXLS(Object Sender, EventArgs e)
        {
            ShowCrossTable_type0(2);
        }
        private void ShowCrossTable_type0(int int_type)
        {
            System.Data.DataTable orig_dt = GetCrossTable_type(int_type);
            System.Data.DataTable dt = new System.Data.DataTable();
            String sql = "select tostaf_ref,formtype,TONAME from stafeval_grid group by tostaf_ref,formtype;";
            System.Data.Common.DbDataReader dr = ESData.GetInst.Reader(sql);
            dt.Load(dr);
            System.Text.Encoding enc = System.Text.Encoding.GetEncoding("us-ascii");
            for (byte i = 65; i < 91; i++)
            {
                byte[] bytes = { i };
                dt.Columns.Add(enc.GetString(bytes));
            }

            dt.Columns.Add("total");
            dt.Columns.Add("cnt");
            dt.Columns.Add("formtype_desc");
            for (byte i = 65; i < 91; i++)
            {
                byte[] bytes = { i };
                dt.Columns.Add("Q" + enc.GetString(bytes));
            }
            foreach (DataRow row in dt.Rows)
            {
                DataRow[] result = orig_dt.Select(String.Format("tostaf_ref='{0}' and formtype='{1}'  ", row[0], row[1]));
                row["cnt"] = result.Length;
                decimal[] dec_temp = new decimal[27];
                foreach (DataRow sr in result)
                {
                    for (byte i = 65; i < 91; i++)
                    {
                        byte[] bytes = { i }; String f = enc.GetString(bytes);
                        if (!sr.IsNull(f) && sr[f].ToString().Length > 0)
                            dec_temp[i - 65] += decimal.Parse(sr[f].ToString());
                    }
                    dec_temp[26] += decimal.Parse(sr["total"].ToString());
                    if (row["formtype_desc"].ToString().Length < 2)
                    {
                        row["formtype_desc"] = sr["formtype_desc"].ToString();
                        for (byte i = 65; i < 91; i++)
                        {
                            byte[] bytes = { i }; String f = enc.GetString(bytes);
                            row["Q" + f] = sr["Q" + f].ToString();
                        }
                    }
                }
                for (byte i = 65; i < 91; i++)
                {
                    byte[] bytes = { i }; String f = enc.GetString(bytes);
                    if (dec_temp[i - 65] > 0) { row[f] = dec_temp[i - 65] / result.Length; }
                }
                row["total"] = dec_temp[26] / result.Length;
            }


            //ES_FORMS.iDTFromGrid idt = new ES_FORMS.iDTFromGrid(dt);
            //ES_FORMS.FormDataGrid_Cmd fdg = new FormDataGrid_Cmd(idt, null, BindingListOptions.AllowModifyNo);
            //fdg.MdiParent = this;
            //fdg.Show();
        }
        public static System.Data.DataTable GetCrossTable_type(int int_type)
        {
            System.Collections.Hashtable formdesc = new System.Collections.Hashtable();
            System.Collections.Hashtable qizrate = new System.Collections.Hashtable();

            System.Collections.Hashtable qiz_desc = new System.Collections.Hashtable();
            String qizratesql = "select key0,rate,stype,formtype,formname from stafeval_gridqiz where key0 like '%e';";
            System.Data.Common.DbDataReader qizratedr = ESData.GetInst.Reader(qizratesql);
            while (qizratedr.Read())
            {
                qizrate.Add(qizratedr.GetString(0).Substring(0, 3), qizratedr.GetInt32(1));
                qiz_desc.Add(qizratedr.GetString(0).Substring(0, 3), qizratedr.GetString(2));
                if (!formdesc.ContainsKey(qizratedr.GetString(3)))
                {
                    formdesc.Add(qizratedr.GetString(3), qizratedr.GetString(4));
                }
            }
            qizratedr.Close();
            qizratedr.Dispose();
            System.Data.DataTable dt = new System.Data.DataTable();
            String sql = "select staf_ref,c_name,tostaf_ref,toname,formtype from stafeval_grid group by staf_ref,tostaf_ref,formtype,c_name,toname;";
            System.Data.Common.DbDataReader dr = ESData.GetInst.Reader(sql);
            dt.Load(dr);
            System.Text.Encoding enc = System.Text.Encoding.GetEncoding("us-ascii");
            for (byte i = 65; i < 91; i++)
            {
                byte[] bytes = { i };
                dt.Columns.Add(enc.GetString(bytes));
            }
            dt.Columns.Add("total");
            dt.Columns.Add("cnt");
            dt.Columns.Add("formtype_desc");
            for (byte i = 65; i < 91; i++)
            {
                byte[] bytes = { i };
                dt.Columns.Add("Q" + enc.GetString(bytes));
            }

            String sql0 = "select staf_ref,tostaf_ref,formtype,qiz,ans0 from stafeval_grid ;";
            System.Data.Common.DbDataReader dr0 = ESData.GetInst.Reader(sql0);
            while (dr0.Read())
            {
                byte b = (byte)(64 + dr0.GetInt32(3));
                byte[] bytes = { b };
                String f = enc.GetString(bytes);
                DataRow[] result = dt.Select(String.Format("staf_ref='{0}' and tostaf_ref='{1}' and formtype='{2}'  ", dr0[0], dr0[1], dr0[2]));
                foreach (DataRow row in result)
                {
                    row["formtype_desc"] = formdesc[row["formtype"].ToString().ToLower()];
                    String key0 = dr0.GetString(2) + f;
                    if (qizrate.ContainsKey(key0.ToLower()))
                    {
                        row["Q" + f] = qiz_desc[key0.ToLower()].ToString();
                    }
                    //Console.WriteLine("debug:" + dr0.GetString(0) + dr0.GetString(1));
                    //if (dr0.IsDBNull(dr0.GetOrdinal("ans0")) || dr0["ans0"].ToString().Equals("")) 
                    if (dr0.IsDBNull(4) || dr0["ans0"].ToString().Equals(""))
                    {

                    }
                    else if (int_type == 1 && dr0["ans0"].ToString().Length > 0)
                    {
                        int v = 0;
                        if (int.TryParse(dr0["ans0"].ToString()[0].ToString(), out v))
                        {
                            row[f] = v;
                        }
                    }
                    else if (int_type == 2 && dr0["ans0"].ToString().Length > 0)
                    {
                        int v = 0;
                        if (int.TryParse(dr0["ans0"].ToString()[0].ToString(), out v))
                        {
                            int r = 1;
                            if (qizrate.ContainsKey(key0.ToLower()))
                            {
                                r = (int)qizrate[key0.ToLower()];
                            }
                            else { Console.WriteLine("error:" + dr0.GetString(0) + dr0.GetString(1) + key0); }
                            row[f] = v * r / 5.0;
                        }
                    }
                    else
                    {
                        row[f] = dr0["ans0"];
                    }
                }

            }
            dr.Close();
            dr.Dispose();
            if (int_type == 1 || int_type == 2)
            {
                foreach (DataRow row in dt.Rows)
                {
                    decimal total = 0.0M;
                    for (int i = 0; i < 26; i++)
                    {
                        byte b = (byte)(65 + i);
                        byte[] bytes = { b };
                        String f = enc.GetString(bytes);
                        if (!row.IsNull(f) && row[f].ToString().Length > 0)
                        {
                            try
                            {
                                total += decimal.Parse(row[f].ToString());
                            }
                            catch (Exception e1) { Console.WriteLine(row[f].ToString() + e1.Message); }
                        }
                    }
                    row["total"] = total;
                }
            }
            return dt;
        }

    }
}
