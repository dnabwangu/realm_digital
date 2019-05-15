using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Net;
using Newtonsoft.Json;
using Tara_app.Services;
using System.IO;

namespace Tara_app
{
    public partial class ThisAddIn
    {
        public class MyWebClient : WebClient
        {
          

            public MyWebClient()
            {
                var authHeader = AuthHandler.Instance.GetAuthHeader();
                Headers.Add("Authorization", authHeader);
            }
            protected override WebRequest GetWebRequest(Uri uri)
            {
                WebRequest w = base.GetWebRequest(uri);
                w.Timeout = 10 * 60 * 20000;
                return w;
            }
        }

        public string Get_base_url()
        {
            //string url = "https://api.ai-spark.com/v1/";
            string url = "http://127.0.0.1:5000/v1/";
            return (url);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string server = Get_base_url();
            Excel.Application TaraApp = (Excel.Application)Globals.ThisAddIn.Application;

            string result_path = TaraApp.Application.UserLibraryPath + "\\tara\\";
            bool exists = System.IO.Directory.Exists(result_path);
            if (!exists)
                System.IO.Directory.CreateDirectory(result_path);
            // delete all files in results directory PUT BACK WITH MODEL GOVERNANCE

            System.IO.DirectoryInfo di = new DirectoryInfo(result_path);
            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }
        public string Get_used_range(Array  array)
        {
            int consecutive_null_rows = 0;
            int consecutive_null_cols = 0;
            int row_count = 1;
            int average_cell_count = 0;
            int sum_cell_count = 0;
            int max_cell_count = 0;

            for (int array_row = 1; array_row < array.GetLength(0)+1; array_row++)
            {
                int cell_count = 0;
                for (int array_col = 1; array_col < array.GetLength(1); array_col++)
                {
                    string cell = null;
                    try
                    {
                        cell = Convert.ToString(array.GetValue(array_row, array_col));
                    }
                    catch { cell = null; }
                        
                        
                   
                    if (cell == null || cell == "")
                    {
                        consecutive_null_cols += 1;
                    }
                    else
                    {
                        cell_count =array_col;
                        consecutive_null_cols = 0;
                    }
                    if (consecutive_null_cols == 100) { break; }
                }

                if (cell_count < average_cell_count/4) {
                    row_count += 1;
                    consecutive_null_rows += 1;  }
                else
                {
                    row_count += 1;
                    consecutive_null_rows = 0;
                    sum_cell_count += cell_count;
                    average_cell_count = sum_cell_count / row_count;
                    if (cell_count > max_cell_count) { max_cell_count = cell_count; }
                }
                if (consecutive_null_rows == 10) { break; }
            }
            row_count = row_count - consecutive_null_rows-1;
            return row_count + "|"+max_cell_count;
        }
        public Tuple<string, int> Get_Data()
        {
            Excel.Worksheet workSheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            if (workSheet.AutoFilter != null) { workSheet.AutoFilterMode = false;  }
            System.Array myvalues;
            string JSON_data;
            int iTotalRows;
            try
            {
                myvalues = (System.Array)workSheet.UsedRange.Cells.Value;
                string[] coords = Get_used_range(myvalues).Split('|');
                iTotalRows = Convert.ToInt32(coords[0]);
                int iTotalColumns = Convert.ToInt32(coords[1]) + 1;

                Excel.Range finalCell = workSheet.Cells[iTotalRows, iTotalColumns];
                Excel.Range range = workSheet.get_Range("A1", finalCell);
                myvalues = (System.Array)range.Cells.Value;
                if (iTotalColumns < 10)
                {
                    JSON_data = "no data";
                    iTotalRows = 0;
                }
                else
                {
                    JSON_data = JsonConvert.SerializeObject(myvalues);

                }

            }
            catch{
                JSON_data = "no data";
                iTotalRows = 0;
            }
            return Tuple.Create(JSON_data, iTotalRows);

        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
