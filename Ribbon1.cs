using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelDataReader;
using Newtonsoft.Json;
using Tara_app.Services;

namespace Tara_app
{
    public partial class AI_Spark
    {
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            if (AuthHandler.Instance.IsUserAuthenticated() == true)
            {
                ShowAuthenticatedLayout();
            } else
            {
                ShowUnAuthenticatedLayout();
            }
        }

        public void ShowUnAuthenticatedLayout()
        {
            data.Visible = false;
            comparison.Visible = false;
            //group3.Visible = false;
            view.Visible = false;

            logoutButton.Visible = false;
            userLabel.Visible = false;
            loginButton.Visible = true;
        }

        public void ShowAuthenticatedLayout()
        {
            LoadData();
            data.Visible = true;
            view.Visible = true;
            comparison.Visible = true;

            loginButton.Visible = false;
            logoutButton.Visible = true;
            userLabel.Visible = true;

            userLabel.Label = AuthHandler.Instance.GetAuthUser();
        }

        public void LoadData()
        {
            Load_compare_drop_down();
            //Load_deal_drop_down();
        }

        public Array Get_intex_selections(string type)
        {
            using (WebClient client = new ThisAddIn.MyWebClient())
            {
                string responsebody;
                byte[] responsebytes;
                string selected_deal = "";
                if (type != "deals") {
                    selected_deal = type;
                    type = "yymm";
                }

                var reqparm = new System.Collections.Specialized.NameValueCollection
                {
                    { "type", type }
                };

                reqparm.Add("selected_deal", selected_deal);

                responsebytes = client.UploadValues(Globals.ThisAddIn.Get_base_url() + "get_intex_selections", "POST", reqparm);
                responsebody = Encoding.UTF8.GetString(responsebytes);
                responsebody = responsebody.Replace("[", "").Replace("]", "").Replace("\"","").Replace(" ","");

                Array list_ = responsebody.Split(',');
            
                return list_;
            }
        }
      
        private void Load_compare_drop_down()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "vintage";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "spark_property_type";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "spark_loan_size";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "spark_property_type-- vintage";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "spark_property_type-- spark_loan_size";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "base_year-- base_quarter";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "vintage-- base_year-- base_quarter";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "spark_property_type-- base_year-- base_quarter";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "spark_property_type-- vintage-- base_year-- base_quarter";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "spark_property_type-- spark_loan_size-- base_year-- base_quarter";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "spark_property_type-- spark_loan_size-- base_year-- base_quarter";
            compare_dropdown.Items.Add(temp_dropdown_item);
            temp_dropdown_item = this.Factory.CreateRibbonDropDownItem();
            temp_dropdown_item.Label = "spark_loan_size-- base_year-- base_quarter";
            compare_dropdown.Items.Add(temp_dropdown_item);
   
        }
     
        public void Apply_formats()
        {
            Excel.Application TaraApp = (Excel.Application)Globals.ThisAddIn.Application;
            Excel.Worksheet tara_sheet = (Excel.Worksheet)TaraApp.Worksheets["tara_sheet"];

            Excel.Range arange = tara_sheet.UsedRange;
            //arange.Columns.ColumnWidth = 15;
            arange.Cells.Font.Size = 8;
            TaraApp.Windows.Application.ActiveWindow.DisplayGridlines = false;
            tara_sheet.Rows[1].EntireRow.Font.Bold = true;
            tara_sheet.Range["d2"].Select();
            TaraApp.ActiveWindow.FreezePanes = true;

            Excel.Range uw_range = tara_sheet.UsedRange.Columns["G:I", Type.Missing];
            uw_range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            uw_range.Borders.Weight = Excel.XlBorderWeight.xlThin;

            int eTotalColumns = tara_sheet.UsedRange.Columns.Count;
            int col_ = 1;

            tara_sheet.Rows[1].AutoFilter();
            arange.Columns.AutoFit();

            for (col_ =1; col_<=20; col_++)
            {
                Excel.Range Prng = tara_sheet.Columns[col_];
                Prng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                string title = tara_sheet.Cells[1, col_].value;
                if (title != null)
                {
                    if (title.Contains("loss_in_"))
                    {
                        object mis = Type.Missing;
                        Excel.Databar data_bar = Prng.FormatConditions.AddDatabar();
                        data_bar.MaxPoint.Modify(XlConditionValueTypes.xlConditionValueNumber, 0.3);
                        data_bar.MinPoint.Modify(XlConditionValueTypes.xlConditionValueNumber, 0.0);
                    }
                    if (title == "loss_expectation" || title == "loss_expectation_recession" || title == "tara_rating" || title == "property_name" || title == "tara_rating" || title.Contains("loss_in_") || title == "key_drivers" || title == "comparison" || title == "latent_potential")
                    {
                        Prng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        Prng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue);
                    }
                    if (title == "comparison" || title == "impact")
                    {
                        Prng.NumberFormat = "#,###.##";
                        Prng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue);

                    }
                    if (title == "lgd" || title == "actual_loss_in_12" || title == "actual_loss_in_36" || title == "loss_pct" || title == "loss_in_12" || title == "loss_in_36" || title == "loss_in_lifetime" || title == "loss_expectation" || title == "loss_expectation_recession")
                    {
                        Prng.NumberFormat = "0.0%";
                        if (title == "lgd" || title == "loss_pct" || title == "actual_loss_in_12" || title == "actual_loss_in_36" || title == "loss_in_lifetime")
                        {
                            Prng.Hidden = true;
                        }
                        //Prng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue);

                    }
                    if (title == "key_drivers")
                    {
                        Prng.Columns.ColumnWidth = 25;
                    }
                    //if (title == "loss_pct" || title == "actual_lgd" || title == "actual_loss_in_12" || title == "actual_loss_in_36" || title == "actual_loss_in_lifetime")
                    //{
                    //    Prng.Hidden = true;
                    //}
                }
            }
            tara_sheet.EnableOutlining = true;
            //Excel.Range group_range = tara_sheet.Range["B:D"];
            //group_range.Columns.Group();
            //group_range.Columns.OutlineLevel = 2;
            Excel.Range group_range_comp = tara_sheet.Range["D:F"];
            group_range_comp.Columns.Group();
            group_range_comp.Columns.OutlineLevel = 2;
            Excel.Range group_range_admin = tara_sheet.Range["M:R"];
            group_range_admin.Columns.Group();
            group_range_admin.Columns.OutlineLevel = 4;
            Excel.Range group_range_pct = tara_sheet.Range["AD:AL"];
            group_range_pct.Columns.Group();
            group_range_pct.Columns.OutlineLevel = 4;

            tara_sheet.Outline.ShowLevels(0, 1);
        }

        private void Get_illumination()
        {
            Excel.Application TaraApp = (Excel.Application)Globals.ThisAddIn.Application;
            var reqparm = new System.Collections.Specialized.NameValueCollection();

            Excel.Worksheet tara_sheet = (Excel.Worksheet)TaraApp.Worksheets["tara_sheet"];
            tara_sheet.UsedRange.Columns[9,Type.Missing].Interior.ColorIndex = 0;
            string JSON_data = Globals.ThisAddIn.Get_Data().Item1;
            using (WebClient client = new ThisAddIn.MyWebClient())
            {
                string server = Globals.ThisAddIn.Get_base_url();
                string responsebody = "";

                reqparm.Add("data", JSON_data);

                byte[] responsebytes = client.UploadValues(server + "illuminate", "POST", reqparm);
                responsebody = Encoding.UTF8.GetString(responsebytes);
                responsebody = responsebody.Replace("|", "--").Replace("][", "|").Replace("[", "").Replace("]", "");

                Excel.Range pRng = tara_sheet.Range["A1"];
                Array responses = responsebody.Split('|');
                Double sensitivity;
                int eTotalColumns = tara_sheet.UsedRange.Columns.Count;
                int base_loss_col = 1;
                foreach (string c in responses)
                {
                    int place_col = 1;
                    string response_header = c.Split(',')[0].Trim(' ').Trim('"');

                    for (int e_col = place_col; e_col <= eTotalColumns; e_col++)
                    {
                        if (tara_sheet.Cells[1, e_col].Value2 != null)
                        {
                            string existing_header = tara_sheet.Cells[1, e_col].Value.ToString().Trim(' ').Trim('"');
                            if (existing_header == "loss_expectation")
                            {
                                base_loss_col = e_col;
                            }

                            else
                            {
                                string r_h = response_header.ToLower().Replace(" ", "").Replace(")", "").Replace("(", "").Replace(":", "").Replace(";", "").Replace("-", "").Replace("'", "").Replace("]", "").Replace("[", "").Trim('"').Replace("\"", "").Replace(",", "");
                                string e_h = existing_header.ToLower().Replace(" ", "").Replace(")", "").Replace("(", "").Replace(":", "").Replace(";", "").Replace("-", "").Replace("'", "").Replace("]", "").Replace("[", "").Trim('"').Replace("\"", "").Replace(",", "");
                                if (r_h == "keydrivers" & e_h == "key_drivers")
                                {

                                    int place_row = 1;

                                    for (int row_c = 1; row_c < c.Split(',').Length; row_c += 1)
                                    {
                                        string v = c.Split(',')[row_c].Replace("\\\\r\\\\n","\r\n").Replace("\\", "").Replace("\"", "").Replace("--", "|").Replace("*","-");
                                        tara_sheet.Cells[row_c + place_row, e_col].Value2 = v;

                                    }
                                    tara_sheet.UsedRange.Columns[e_col].WrapText = false;
                                }
                                else if (r_h == "latent_potential")
                                {
                                    if (e_h == r_h)
                                    {
                                        int place_row = 1;

                                        for (int row_c = 1; row_c < c.Split(',').Length; row_c += 1)
                                        {
                                            Double v = Convert.ToDouble(c.Split(',')[row_c]);
                                            string v_desc = "none";
                                            if (v > .04) { v_desc = "volatile"; }
                                            else if (v > .03) { v_desc = "high"; }
                                            else if (v > .015) { v_desc = "elevated"; }
                                            else if (v > 0.000) { v_desc = "moderate"; }
                                            else { v_desc = "low"; }
                                            tara_sheet.Cells[row_c + place_row, e_col].Value2 = v_desc;

                                        }
                                    }
                                    else if (e_h == "tara_rating")
                                    {
                                        int place_row = 1;

                                        for (int row_c = 1; row_c < c.Split(',').Length; row_c += 1)
                                        {
                                            Double v = Convert.ToDouble(c.Split(',')[row_c]);
                                            int c_index = 1;
                                            sensitivity = 0;
                                            c_index = 1;
                                            if (v > .4) { c_index = 3; sensitivity = .9; }
                                            else if (v > .2) { c_index = 3; sensitivity = .25; }
                                            else if (v > -.1) { c_index = 3; sensitivity = 0; }
                                            else if (v > -.2) { c_index = 3; sensitivity = -0.25; }
                                            else { c_index = 3; sensitivity = -.9; }
                                            //sensitivity = Math.Max(0.2, Math.Min(1,Math.Abs(v)));
                                            Excel.Range iRng = tara_sheet.Cells[row_c + place_row, e_col];
                                            if (!String.IsNullOrEmpty(Convert.ToString(iRng.Cells.Value2)))
                                            {
                                                iRng.Font.Color = c_index;
                                                iRng.Font.TintAndShade = sensitivity;
                                            }
                                        }
                                    }

                                }
                                else if (e_h == r_h)
                                {
                                    int place_row = 1;

                                    for (int row_c = 1; row_c < c.Split(',').Length; row_c += 1)
                                    {
                                        Double v = Convert.ToDouble(c.Split(',')[row_c]);
                                        int c_index = 3;
                                        if (v > 0) { c_index = 23; }
                                        sensitivity = Math.Max(0, Math.Min(1, 1 - Math.Abs(v)));
                                        Excel.Range iRng = tara_sheet.Cells[row_c + place_row, e_col];
                                        if (!String.IsNullOrEmpty(Convert.ToString(iRng.Cells.Value2)))
                                        {
                                            iRng.Interior.ColorIndex = c_index;
                                            iRng.Interior.TintAndShade = sensitivity;
                                        }

                                    }
                                    e_col = eTotalColumns;
                                }
                            }
                        }
                    }
                }
            }
            Apply_formats();
            tara_sheet.Outline.ShowLevels(3, 3);

        }

        public List<string> Get_normalized_deals(string FolderPath, string result_file)
        {
            List<string> deal_names = new List<string>();
            if (result_file != "_back_test_data")
            {
                if (!File.Exists(FolderPath + "/" + result_file + ".csv"))
                {
                    File.CreateText(FolderPath + "/" + result_file + ".csv");
                }
                else
                {
                    using (var reader = new StreamReader(FolderPath + "/" + result_file + ".csv"))
                    {
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(',');
                            var value = values[2].Replace("\"", "").Trim();
                            if (deal_names.Contains(value))
                            {
                                //do nothing
                            }
                            else
                            {
                                deal_names.Add(value);
                            }

                        }
                    }
                }
            }
            return deal_names;
        }
        public void NormalizeFiles(string FolderPath, string result_file, string d_date)
        {
            string[] files = Directory.GetFiles(FolderPath);
            List<string> deal_names = new List<string>();
            deal_names = Get_normalized_deals(FolderPath, result_file);

            foreach (string f in files)
            {
                Globals.ThisAddIn.Application.ActiveCell.Value = f;
                string file_path = Path.Combine(FolderPath, f);
                FileStream stream = File.Open(file_path, FileMode.Open, FileAccess.Read);
                System.Data.DataSet result = null;
                var file_name = file_path.Split('\\').Last().Replace(".xlsx", "").Replace(".xls", "").Replace(".csv", "").Replace(".CSV", "").Replace(".XLSX", "").Replace(".XLS", "");
                if (deal_names.Contains(file_name) || file_name.Contains("_back_test_ratings") || file_name.Contains("_back_test_data") || file_name.Contains("_tara_setup_tapes")) { } //do nothing
                else
                {
                    try
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                        {
                            result = reader.AsDataSet();
                        }
                    }
                    catch
                    {
                        try
                        {
                            using (IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(stream))
                            {
                                result = reader.AsDataSet();

                            }
                        }
                        catch
                        {
                            Console.WriteLine(result);
                        }
                    }
                    if (result != null)
                    {
                        string server = Globals.ThisAddIn.Get_base_url();
                        using (WebClient client = new WebClient())
                        {
                            string responsebody = "";
                            var reqparm = new System.Collections.Specialized.NameValueCollection();
                            string JSON_data = JsonConvert.SerializeObject(result);
                            string deal_name = stream.Name;
                            reqparm.Add("param1", JSON_data);
                            reqparm.Add("param2", deal_name);
                            reqparm.Add("get_sup", "yes");
                            reqparm.Add("backtest", "yes");
                            reqparm.Add("d_date", d_date);
                            reqparm.Add("irp", "yes");

                            byte[] responsebytes = client.UploadValues(server + "create_standard_data_sheet", "POST", reqparm);
                            responsebody = Encoding.UTF8.GetString(responsebytes);

                            responsebody = responsebody.Replace("}{", "|").Replace("}", "").Replace("[", "").Replace("{", "");
                            Array raw_loan_data = responsebody.Split('|');
                            string out_f = result_file + ".csv";
                            string written_file = Response_to_csv(raw_loan_data, FolderPath, out_f, true, d_date, true);

                        }

                    }
                }
                stream.Close();

            }
        }

        private string ConvertToCsvCell(string value)
        {
            var mustQuote = value.Any(x => x == ',' || x == '\"' || x == '\r' || x == '\n');
            if (!mustQuote)
            {
                return value;
            }
            value = value.Replace("\"", "").Trim();
            return string.Format("\"{0}\"", value);

        }

        public string Response_to_csv(Array raw_loan_data, string folder_path, string f, Boolean overwrite, string d_date, Boolean backtest)
        {
            string file_path = Path.Combine(folder_path, f);
            int row_c = 0;
            var sb = new StringBuilder();

            foreach (string raw_loan in raw_loan_data)
            {

                string row = raw_loan.Replace("\",", "|").Replace(",\"", "|");
                string[] cells = row.Split('|');
                int col_c = 0;

                if (row_c == 0)
                {
                    List<string> csvList_header = new List<string>();
                    col_c = 0;
                    foreach (string c in cells)
                    {
                        string[] key_value = c.Split(':');
                        csvList_header.Add(ConvertToCsvCell(key_value[0].Replace("\"", "")));
                        col_c += 1;
                    }
                    sb.AppendLine(string.Join(",", csvList_header));
                }

                List<string> csvList = new List<string>();
                col_c = 0;
                foreach (string c in cells)
                {
                    string[] key_value = c.Split(':');
                    csvList.Add(ConvertToCsvCell(key_value[1]));
                    col_c += 1;
                }
                if (!csvList.Any()) { }
                else
                {
                    sb.AppendLine(string.Join(",", csvList));
                }
                
                row_c += 1;
            }
            System.IO.StreamWriter streamWriter;
            string file_assignment = "none";
            int attempt = 0;
            while (file_assignment != "done")
            {
                attempt += 1;
                if (System.IO.File.Exists(file_path.Replace(".csv", "") + "_" + Convert.ToString(attempt) + ".csv")) {
                    // do nothing
                }
                else { file_assignment = "done";}
              
            }
            string written_file = file_path.Replace(".csv", "") + "_" + Convert.ToString(attempt) + ".csv";
            bool exists = System.IO.Directory.Exists(folder_path);
            if (!exists)
                System.IO.Directory.CreateDirectory(folder_path);
            streamWriter = new System.IO.StreamWriter(file_path.Replace(".csv", "") + "_" + Convert.ToString(attempt) + ".csv", overwrite);
            streamWriter.Write(sb.ToString());
            streamWriter.Close();
            return written_file;
        }
  
        private void Ddates_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
        }
        private void Run_backtest(string folder_path, string file_name)
        {

            String line = String.Empty;
            String[] cells = null;
            System.IO.StreamReader file = new System.IO.StreamReader(folder_path + "/" + file_name + "_1.csv");
            int row = 0;
            int line_length = 0;
            List<string[]> myvalues = new List<string[]>();
            while ((line = file.ReadLine()) != null)
            {
                cells = line.Replace("\",\"", "\"|\"").Split('|');
                if (row == 0) { line_length = cells.GetLength(0); }
                if (cells.GetLength(0) == line_length)
                {
                    for (int c = 0; c < cells.GetLength(0); c++)
                    {
                        cells[c] = cells[c].Replace("\"", "");
                    }
                    myvalues.Add(cells);
                }
                row += 1;
            }
            String[][] contents = myvalues.ToArray();
            if (contents != null)
            {
                string server = Globals.ThisAddIn.Get_base_url();
                using (WebClient client = new WebClient())
                {
                    string responsebody = "";
                    var reqparm = new System.Collections.Specialized.NameValueCollection();
                    string JSON_data = JsonConvert.SerializeObject(contents);
                    reqparm.Add("data", JSON_data);
                    reqparm.Add("d_date", "");
                    reqparm.Add("get_sup", "no");


                    byte[] responsebytes = client.UploadValues(server + "ask_tara", "POST", reqparm);
                    responsebody = Encoding.UTF8.GetString(responsebytes);
                    responsebody = responsebody.Replace("[", "").Replace("]", "");
                    List<string> loss = responsebody.Split(',').ToList();
                    loss.Insert(0, "Tara E(loss)");
                    List<string[]> response = new List<string[]>();
                    List<string> rows = new List<string>();

                    for (int i = 0; i < myvalues.Count(); i++)
                    {
                        string add = "yes";
                        if (myvalues[i][5].Contains("issuance") && myvalues[i][5] != "issuance_at_" + myvalues[i][32])
                        {
                            add = "no";
                        }
                        if (add == "yes")
                        {
                            rows.Add(Convert.ToString(loss[i]));
                            for (int c = 0; c < myvalues[i].Count(); c++)
                            {
                                rows.Add(myvalues[i][c]);
                            }
                            response.Add(rows.ToArray());
                        }
                    }
                    string out_f = "_back_test_ratings.csv";
                    string file_path = Path.Combine(folder_path, out_f);
                    int row_c = 0;
                    var sb = new StringBuilder();

                    foreach (string[] raw_loan in response)
                    {
                        List<string> csvList = new List<string>();
                        foreach (string cell in raw_loan)
                        {
                            csvList.Add(ConvertToCsvCell(cell));
                        }
                        if (!csvList.Any()) { }
                        else
                        {
                            sb.AppendLine(string.Join(",", csvList));
                        }
                        row_c += 1;
                    }
                    System.IO.StreamWriter streamWriter;
                    streamWriter = new System.IO.StreamWriter(file_path, true);
                    streamWriter.WriteLine(sb.ToString());
                    streamWriter.Close();
                }
            }
            file.Close();
        }
        private void Backtesting_Click(object sender, RibbonControlEventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string file_name = "_back_test_data";
                string folder_path = folderBrowserDialog1.SelectedPath;
                for (int year = 2005; year <= 2006; year++)
                {
                    for (int m = 1; m <= 3; m++)
                    {
                        string month = m.ToString("D2");
                        string d_date = Convert.ToString(year) + "-" + month;
                        NormalizeFiles(folder_path, file_name, d_date);
                    }
                }
                Run_backtest(folder_path, file_name);
            }
        }


       

        private void Compare_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

       
        private void Compare_dropdown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void DropDown3_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void DropDown4_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void DropDown5_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }
        private void Tara_ratings()
        {

            Excel.Application TaraApp = (Excel.Application)Globals.ThisAddIn.Application;
            Excel.Worksheet tara_sheet;
            string tara_sheet_ = "tara_sheet";
            try
            {
                tara_sheet = (Excel.Worksheet)TaraApp.Worksheets["tara_sheet"];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                tara_sheet_ = "no";
            }
            if (tara_sheet_ != "no")
            {
                tara_sheet = (Excel.Worksheet)TaraApp.Worksheets[tara_sheet_];
                using (WebClient client = new ThisAddIn.MyWebClient()) //client = new WebClient())
                {

                    string server = Globals.ThisAddIn.Get_base_url();
                    string responsebody = "";

                    var reqparm = new System.Collections.Specialized.NameValueCollection();
                    var _data = Globals.ThisAddIn.Get_Data();

                    reqparm.Add("data", _data.Item1);
                    reqparm.Add("d_date", null);
                    reqparm.Add("get_sup", "no");

                    byte[] responsebytes = client.UploadValues(server + "ask_tara", "POST", reqparm);
                    responsebody = Encoding.UTF8.GetString(responsebytes);
                    //responsebody = responsebody.Replace("},{", "|").Replace("[", "").Replace("]", "").Replace("\"", "").Replace("}", "").Replace("{", "");
                    responsebody = responsebody.Replace("|", "--").Replace("____change___", "|");
                    String[] body_change = responsebody.Split('|');

                    responsebody = body_change[0];
                    String responsechange = body_change[1];
                    responsebody = responsebody.Replace("], \"", "|").Replace("[", "").Replace("]", "").Replace("{", "").Replace("}", "");
                    responsechange = responsechange.Replace("], \"", "|").Replace("[", "").Replace("]", "");

                    Excel.Range pRng = tara_sheet.Range["A1"];

                    Array responses = responsebody.Split('|');
                    String[] changes = responsechange.Split('|');
                    string title = "";


                    if (Convert.ToString(responses.GetValue(0)) != "\"could not find data\"")
                    {
                        int col_count = 0;
                        foreach (string r in responses)
                        {

                            // if value FALSE is in the corresponding changes array then dont' run
                            if (changes[col_count].Contains("false")) {
                                //column has an update proceed
                            
                                var column = r.Split(':')[0].Replace("\"", "");
                                int place_col = 0;


                                for (int place_col_ = 1; place_col_ <= responses.Length; place_col_ += 1)
                                {
                                    title = tara_sheet.Cells[1, place_col_].Value2;
                                    if (column == title) { place_col = place_col_;  break; }
                                }
                                if (place_col != 0 & title != "")
                                {
                                    int row_count = r.Split(':')[1].Replace("\", ", "|").Split('|').Length + 1;
                                    for (int place_row = 2; place_row <= row_count; place_row += 1)// can change this to only loop those fields with a FALSE value in change
                                    {
                                        if (changes[col_count].Split(':')[1].Split(',')[place_row - 2].Contains("false"))
                                        { // proceed found a change 
                                            var cell = r.Split(':')[1].Replace("\", ", "|").Split('|')[place_row - 2].Replace("\"", "").Trim();

                                            tara_sheet.Cells[place_row, place_col].Value2 = cell;
                                            if (title == "fy_ncf" | title == "mr_occupancy" | title == "stress") { 
                                                tara_sheet.Cells[place_row,place_col].Interior.ColorIndex = 35;
                                            }
                                        }
                                    }
                                }
                            }
                            col_count += 1;
                        }
                    }
                }
                Apply_formats();

            }
        }
        private void Ask_tara__Click(object sender, RibbonControlEventArgs e)
        {
            string tara_sheet_ = "yes";
            Excel.Application TaraApp = (Excel.Application)Globals.ThisAddIn.Application;
            try
            {
                Excel.Worksheet tara_sheet = (Excel.Worksheet)TaraApp.Worksheets["tara_sheet"];
                tara_sheet_ = "tara_sheet";
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                tara_sheet_ = "no"; 
            }

            if (tara_sheet_ == "no")
            {
                //do nothing 
            }
            else
            {

                Tara_ratings();
            }

        }

        private void Illuminate__Click(object sender, RibbonControlEventArgs e)
        {


            string tara_sheet_ = "yes";
            Excel.Application TaraApp = (Excel.Application)Globals.ThisAddIn.Application;

            try
            {
                Excel.Worksheet tara_sheet = (Excel.Worksheet)TaraApp.Worksheets["tara_sheet"];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                tara_sheet_ = "no";
            }

            if (tara_sheet_ == "no")
            {
                // do nothing

            }
            else
            {
                Get_illumination();
            }
        }

        private void Compare_button_Click(object sender, RibbonControlEventArgs e)
        {
            string tara_sheet_ = "yes";
            var reqparm = new System.Collections.Specialized.NameValueCollection();
            var _data = Globals.ThisAddIn.Get_Data();
            string responsebody;
            byte[] responsebytes;
            Excel.Application TaraApp = (Excel.Application)Globals.ThisAddIn.Application;
            Excel.Worksheet tara_sheet = (Excel.Worksheet)TaraApp.ActiveSheet; //Worksheets["TARA"];
            object mis = Type.Missing;

            if (_data.Item1 == "no data") { tara_sheet_ = "skip"; }
            if (tara_sheet_ != "skip")
            {
                using (WebClient client = new ThisAddIn.MyWebClient())
                {
                    Excel.Worksheet work_sheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

                    reqparm.Add("data", _data.Item1);
                    reqparm.Add("type", compare_dropdown.SelectedItem.Label);

                    responsebytes = client.UploadValues(Globals.ThisAddIn.Get_base_url() + "load_comparison", "POST", reqparm);
                    responsebody = Encoding.UTF8.GetString(responsebytes);
                    responsebody = responsebody.Replace("{", "").Replace("}", "").Replace("\"", "");

                    Array ranks = responsebody.Split(',');
                    int eTotalColumns = tara_sheet.UsedRange.Columns.Count;
                    int comparison_col = 1;
                    string add_comparison = "yes";
                    for (comparison_col = 1; comparison_col <= eTotalColumns; comparison_col++)
                    {
                        string existing_header = tara_sheet.Cells[1, comparison_col].Value.ToString().Trim(' ').Trim('"');
                        if (Convert.ToString(existing_header).Contains("comparison"))
                        {
                            add_comparison = "no";
                            break;
                        }
                    }
                    if (add_comparison =="yes")
                    {
                        comparison_col = 3;
                        Excel.Range Prng = tara_sheet.UsedRange.Columns[comparison_col];
                        Prng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        Prng.NumberFormat = "0%";
                        tara_sheet.Cells[1,comparison_col].Value2 = "comparison";

                    }
                    foreach (string rank in ranks)
                    {
                        int row = Convert.ToInt16(rank.Trim(' ').Split(':')[0]);
                        String rank_value = rank.Trim(' ').Split(':')[1];
                        work_sheet.Cells[row + 2, comparison_col].Value = rank_value;
                        if (rank_value == "") { rank_value = ".5"; }
                        if (rank_value != "")
                        {
                            Double v = Convert.ToDouble(rank_value) - .5;
                            int c_index = 3;
                            if (v < 0) { c_index = 23; }
                            Double sensitivity = Math.Max(0, Math.Min(1, 1 - Math.Abs(v)));
                            Excel.Range iRng = work_sheet.Cells[row + 2, comparison_col];
                            if (!String.IsNullOrEmpty(Convert.ToString(iRng.Cells.Value2)))
                            {
                                iRng.Interior.ColorIndex = c_index;
                                iRng.Interior.TintAndShade = sensitivity;
                            }
                        }
                    }
                }
            }
            tara_sheet.Outline.ShowLevels(3, 3);

        }

        private void LoginButton_Click(object sender, RibbonControlEventArgs e)
        {
            AuthHandler.Instance.ShowAuthForm();
        }

        private void LogoutButton_Click(object sender, RibbonControlEventArgs e)
        {
            AuthHandler.Instance.Logout();
        }
        

        private void view_form_Click(object sender, RibbonControlEventArgs e)
        {
            var intexForm = new intex_data();
            intexForm.Show();
        }
    }
}

