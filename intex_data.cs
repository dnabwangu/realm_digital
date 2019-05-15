using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Comment = Microsoft.Office.Interop.Excel.Comment;
using System.Net;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace Tara_app
{
    public partial class intex_data : Form
    {
        public intex_data()
        {
            InitializeComponent();
        }

        private void intex_data_Load(object sender, EventArgs e)
        {
            deals.Items.Clear();
            Array Files = Globals.Ribbons.Ribbon1.Get_intex_selections("deals");
            foreach (string file in Files)
            {
                if (!file.Contains("_"))
                {
                    deals.Items.Add(file);
                }
            }
        }

        private void deals_SelectedIndexChanged(object sender, EventArgs e)
        {
            string deal = deals.SelectedItem.ToString();
            yymm.ResetText();
            if (deal != "transaction")
            {
                yymm.Items.Clear();
                Array Files = Globals.Ribbons.Ribbon1.Get_intex_selections(deal);
                foreach (string file in Files)
                {
                    if (!file.Contains("_"))
                    {
                        yymm.Items.Add(file);
                    }
                }
            }

        }

        private void yymm_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (deals.SelectedItem != null) {  
                string deal_ = deals.SelectedItem.ToString();
                string yymm_ = yymm.SelectedItem.ToString();
                if (deal_ != "transaction" & yymm_ != "yymm")
                {
                    var reqparm = new System.Collections.Specialized.NameValueCollection();
                    Excel.Application TaraApp = (Excel.Application)Globals.ThisAddIn.Application;

                    string responsebody;
                    byte[] responsebytes;
                    string result_path = TaraApp.Application.UserLibraryPath + "\\tara\\";
                    reqparm.Add("deal_", deal_);
                    reqparm.Add("yymm_", yymm_);

                    using (WebClient client = new ThisAddIn.MyWebClient())
                    {

                        responsebytes = client.UploadValues(Globals.ThisAddIn.Get_base_url() + "load_intex_data", "POST", reqparm);
                        responsebody = Encoding.UTF8.GetString(responsebytes);


                        responsebody = responsebody.Replace("|", "--").Replace("}[", "|");
                        String[] dropdown_body = responsebody.Split('|');

                        responsebody = dropdown_body[1];
                        String drop_downs = dropdown_body[0];
                        responsebody = responsebody.Replace("},{", "|").Replace("}", "").Replace("]", "").Replace("[", "").Replace("'", "").Replace("{", "");
                        drop_downs = drop_downs.Replace("[", "").Replace("]", "");


                        Array raw_loan_data = responsebody.Split('|');


                        Array buckets_ = drop_downs.Replace("'", "").Replace("\", \"", "|").Split('|');

                        string f = deal_ + "_" + yymm_ + ".csv";

                        string written_file = Globals.Ribbons.Ribbon1.Response_to_csv(raw_loan_data, result_path, f, false, yymm_, false);
                        string result_file_path = Path.Combine(result_path, written_file);
                        TaraApp.Workbooks.Open(@result_file_path);

                        Excel.Worksheet TaraSheet = TaraApp.ActiveWorkbook.ActiveSheet;
                        TaraSheet.Name = "tara_sheet";
                        TaraSheet.Activate();;
                        int iTotalRows = TaraSheet.UsedRange.Rows.Count;
                        int col_count = 0;
                        foreach (string bucket_col_val in buckets_)
                        {
                            string bucket = bucket_col_val.Split(':')[1].Replace("\"", "").Trim();

                            col_count += 1;
                            Excel.Range cRng = TaraSheet.Cells[1, col_count];
                            if (bucket != "None" & bucket != "")
                            {
                                //if (cRng.Comment != null) { cRng.ClearComments(); }
                                //Comment comment = cRng.AddComment(bucket);
                                //comment.Shape.Width = 200;
                                //comment.Shape.Height = 100;

                                var drop_list = new System.Collections.Generic.List<string>();
                                drop_list = bucket.Split(',').ToList();
                                int drop_length = bucket.Length;
                                var list_count = drop_list.Count;
                                int section_length;
                                string flatList = "";
                                if (drop_length > 250)
                                {
                                    section_length = 250 / list_count-1;
                                    for (int spark_num = 0; spark_num < list_count; spark_num += 1)
                                    {
                                        flatList = flatList + drop_list[spark_num].Trim().Substring(0, Math.Min(drop_list[spark_num].Trim().Length,section_length))+",";
                                    }
                                }
                                else
                                {
                                    flatList = string.Join(",", drop_list.ToArray());
                                }


                                //Array flat_list = flatList.ToArray();
                                //var flatList = string.Join(("Spark "+Convert.ToString(Enumerable.Range(1, list_count)).ToArray(),",");
                                for (int row_count = 1; row_count <= iTotalRows; row_count += 1)
                                {
                                    if (row_count < 250 & list_count < 40)
                                    {
                                        var drop_cell = TaraSheet.Cells[row_count, col_count];//.Columns[col_count];//:,col_count]#.Cells[row_count, col_count];

                                        drop_cell.Validation.Delete();
                                        drop_cell.Validation.Add(
                                            XlDVType.xlValidateList,
                                            XlDVAlertStyle.xlValidAlertInformation,
                                            XlFormatConditionOperator.xlBetween,
                                            flatList,
                                            Type.Missing);
                                        drop_cell.Validation.IgnoreBlank = false;
                                        drop_cell.Validation.InCellDropdown = true;
                                        drop_cell.Validation.ShowError = false;
                                    }                                }
                            }
                        }
                        //var in_ = System.IO.File.ReadAllText(file_path);
                        // var inputFile = new FileInfo(export_path); // could be .xls or .xlsx too
                        Excel.Workbook tara_book = (TaraApp.ActiveWorkbook);
                        // tara_book.SaveAs(deal_ + "_" + yymm_);
                        Excel.Worksheet tara_sheet = (TaraApp.ActiveSheet);
                        tara_sheet.Name = "tara_sheet"; //deal_ + "_" + yymm_;
                        Globals.Ribbons.Ribbon1.Apply_formats();
                       
                    }
                }
            }
        

        }


    }
}
