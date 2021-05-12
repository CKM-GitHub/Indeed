using System;using System.Collections.Generic;using System.ComponentModel;using System.Data;using System.Data.SqlClient;using System.Diagnostics;using System.Drawing;using System.IO;using System.Linq;using System.Text;using System.Threading;using System.Threading.Tasks;using System.Windows.Forms;using ClosedXML.Excel;using OpenQA.Selenium;using OpenQA.Selenium.Chrome;using OpenQA.Selenium.Interactions;using OpenQA.Selenium.Support.UI;using ExcelDataReader;namespace Indeed{    public partial class Form1 : Form    {        public Form1()        {            InitializeComponent();        }        private void txt_KeyDown(object sender, KeyEventArgs e)        {            TextBox txt = sender as TextBox;            if (e.KeyCode == System.Windows.Forms.Keys.Enter)            {                SelectNextControl(txt, true, true, true, true);            }        }        private void txt_Enter(object sender, EventArgs e)        {            TextBox txt = sender as TextBox;            txt.BackColor = Color.LightCyan;        }        private void txt_Leave(object sender, EventArgs e)        {            TextBox txt = sender as TextBox;            txt.BackColor = SystemColors.Window;        }        private DataTable CreateDatatable()        {            DataTable dt = new DataTable();            dt.Columns.Add("Title");            dt.Columns.Add("Company");            dt.Columns.Add("都道府県");            dt.Columns.Add("Location");            dt.Columns.Add("お問い合わせのURL");            return dt;        }        private void btnBrowse_Click(object sender, EventArgs e)        {            SaveFileDialog sfd = new SaveFileDialog();            sfd.InitialDirectory = @"C:\";            sfd.Filter = "excel files (*.xlsx))|*.xlsx";            sfd.RestoreDirectory = true;            sfd.CheckPathExists = true;            if (sfd.ShowDialog() == DialogResult.OK)            {                txtFilePath.Text = sfd.FileName;            }        }



        #region Search by Key and Location on Indeed and Export to Excel         private void btnRun_Click(object sender, EventArgs e)        {            IWebDriver driver = new ChromeDriver();            try
            {
                this.Cursor = Cursors.WaitCursor;
                DataTable dtFinal = new DataTable();
                DataTable dtResult = new DataTable();
                dtResult.Columns.Add("Title");
                dtResult.Columns.Add("Company");
                dtResult.Columns.Add("都道府県");
                dtResult.Columns.Add("Location");
                dtResult.Columns.Add("お問い合わせのURL");

                int j = 0;
                bool go = true;//to check last page
                bool IsLastPageNo = false;
                int start = 0;//start page = 0,next page = 10,next page = 20.......

                
                int LastPageNo = 0;
                int TotalCount = 0;
                do
                {
                    driver.Navigate().GoToUrl("https://jp.indeed.com/jobs?q=" + txtKeyword.Text + "&l=" + txtLocation.Text + "&limit=50&radius=25&start=" + start.ToString());//searh url with start param
                                                                                                                                                                              // driver.Navigate().GoToUrl("https://jp.indeed.com/%E6%B1%82%E4%BA%BA?q=%E3%83%97%E3%83%AD%E3%82%B0%E3%83%A9%E3%83%9E&l=%E5%8C%97%E6%B5%B7%E9%81%93&limit=50&radius=25");//searh url with start param

                    Thread.Sleep(3000);//wait page load

                    //IList<IWebElement> popup = driver.FindElements(By.ClassName("popover"));
                    //if (popup.Count() > 0)
                    //{
                    //    Actions action = new Actions(driver);
                    //    action.SendKeys(OpenQA.Selenium.Keys.Escape);
                    //}
                    IList<IWebElement> y = driver.FindElements(By.XPath("//*[@id=\"popover-x\"]"));
                    if (y.Count > 0)
                    {
                        y[0].Click();
                    }

                    if(start == 0)
                    {
                        IList<IWebElement> c1 = driver.FindElements(By.XPath("//*[@id=\"searchCountPages\"]"));
                        string[] strarr = c1[0].Text.Split(' ');
                        TotalCount = Convert.ToInt32(strarr[1]);
                    }
                    
                                    
                    //IList<IWebElement> a = driver.FindElements(By.XPath("//*[@aria-current=\"true\"]"));
                    //int CurrentPageNo = Convert.ToInt32(a[0].Text);

                    //if (LastPageNo >= CurrentPageNo)
                    //    IsLastPageNo = true;
                    //else
                    //{
                    //    LastPageNo = CurrentPageNo;

                        
                    //}



                    //get all title,company,location by array
                    IList<IWebElement> arrTitle = driver.FindElements(By.ClassName("title"));//募集内容 
                    IList<IWebElement> arrCompany = driver.FindElements(By.ClassName("company"));//会社名
                                                                                                 //IList<IWebElement> arrPrefecture = driver.FindElements(By.ClassName("prefecture"));//都道府県
                    IList<IWebElement> arrLocation = driver.FindElements(By.ClassName("location"));//所在地

                    for (int i = 0; i < arrTitle.Count; i++)
                    {
                        dtResult.Rows.Add();
                        dtResult.Rows[j]["Title"] = arrTitle[i].Text;
                        dtResult.Rows[j]["Company"] = arrCompany[i].Text;
                        dtResult.Rows[j]["都道府県"] = txtLocation.Text;
                        dtResult.Rows[j]["Location"] = arrLocation[i].Text;
                        dtResult.Rows[j]["お問い合わせのURL"] = string.Empty;
                        j++;
                    }

                    //dtFinal = dtResult;//ktp - to remove  
                    if (dtResult.Rows.Count > 0)
                    {
                        dtFinal = dtResult.AsEnumerable()
                                                         .GroupBy(x => x.Field<string>("Company"))
                                                         .Select(x => x.First())
                                                         .CopyToDataTable();
                    }

                    start += 50;//next page


                    //var CountLi = 0;
                    //IList<IWebElement> arrPaging;
                    //try
                    //{
                    //    arrPaging = driver.FindElements(By.ClassName("pn"));
                    //}
                    //catch
                    //{
                    //    arrPaging = driver.FindElements(By.ClassName("np"));
                    //}
                    //if (!(arrPaging.Count() == 0))
                    //{
                    //    CountLi = arrPaging.Count() + 1;

                    //    var result = driver.FindElement(By.XPath("//*[@id=\"resultsCol\"]/nav/div/ul/li[" + CountLi + "]"));

                    //    var bTagCheck = result.GetAttribute("innerHTML");
                    //    if (bTagCheck.Contains("</b>"))
                    //    {
                    //        IsLastPageNo = true;
                    //    }
                    //}
                    //else IsLastPageNo = true;
                } while (start <= TotalCount);

                string saveFolder = @"C:\Indeed\Search_Result\2021_05_03";
                if (!Directory.Exists(saveFolder))
                {
                    Directory.CreateDirectory(saveFolder);
                }
                SaveFileDialog savedialog = new SaveFileDialog();
                savedialog.Filter = "Excel Files|*.xlsx;";
                savedialog.Title = "Save";
                savedialog.FileName = "_Result";
                savedialog.InitialDirectory = saveFolder;
                savedialog.RestoreDirectory = true;

                if (savedialog.ShowDialog() == DialogResult.OK)
                {
                    if (Path.GetExtension(savedialog.FileName).Contains(".xlsx"))
                    {
                        using (XLWorkbook wb = new XLWorkbook())
                        {
                            wb.Worksheets.Add(dtFinal, "SearchResult");
                            wb.SaveAs(savedialog.FileName);
                        }
                        Process.Start(Path.GetDirectoryName(savedialog.FileName));
                    }
                }
            }            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }            finally
            {
                driver.Quit();
            }                        

            //using (XLWorkbook wb = new XLWorkbook())
            //{
            //    wb.Worksheets.Add(dtResult, "Indeed Search Result");
            //    wb.SaveAs(txtFilePath.Text);

            //    Process.Start(txtFilePath.Text);
            //}
        }




        #endregion
        #region Import Excel Files to DB        private void btnImport_Click(object sender, EventArgs e)        {            DataTable dtImport = new DataTable();            DataTable dt = new DataTable();            OpenFileDialog op = new OpenFileDialog();            op.InitialDirectory = @"C:\Indeed\";            op.Filter = "excel files (*.xlsx))|*.xlsx";            op.RestoreDirectory = true;            op.CheckPathExists = true;            if (op.ShowDialog() == DialogResult.OK)            {                txtFilePath.Text = op.FileName;                Stream st = File.Open(op.FileName, FileMode.Open, FileAccess.Read);                IExcelDataReader reader = ExcelReaderFactory.CreateReader(st);                DataSet ds = reader.AsDataSet(new ExcelDataSetConfiguration()                {                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()                    {                        UseHeaderRow = true                    }                });                dt = ds.Tables[0];                dtImport = CreateDatatable();                for (int i = 0; i < dt.Rows.Count; i++)                {                    dtImport.Rows.Add();                    dtImport.Rows[i]["Title"] = dt.Rows[i]["募集内容"].ToString();                    dtImport.Rows[i]["Company"] = dt.Rows[i]["会社名"].ToString();                    dtImport.Rows[i]["Location"] = dt.Rows[i]["所在地"].ToString();                }                string xml = DataTableToXml(dtImport);                if (MessageBox.Show("Are you sure you want to Import this file?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)                {                    if (Import(xml))                        MessageBox.Show("Import Successfully.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);                }            }        }        private string DataTableToXml(DataTable dt)        {            dt.TableName = "test";            System.IO.StringWriter writer = new System.IO.StringWriter();            dt.WriteXml(writer, XmlWriteMode.WriteSchema, false);            string result = writer.ToString();            return result;        }        private bool Import(string XML)        {            Dictionary<string, ValuePair> dic = new Dictionary<string, ValuePair>            {                {"@xml ", new ValuePair{ value1 = SqlDbType.Xml, value2 = XML } }            };            SqlConnection con = new SqlConnection(@"Data Source=DEVSERVER\SQL2014;Initial Catalog=Indeed;Persist Security Info=True;User ID=sa;Password=admin12345!;");            SqlCommand cmd = new SqlCommand("Insert_Indeed_Data", con);            cmd.CommandType = CommandType.StoredProcedure;            foreach (KeyValuePair<string, ValuePair> pair in dic)            {                ValuePair vpair = pair.Value;                AddParam(cmd, pair.Key, vpair.value1, vpair.value2);            }            con.Open();            cmd.ExecuteNonQuery();            con.Close();            return true;        }        private void AddParam(SqlCommand cmd, string key, SqlDbType dbType, string value)        {            if (string.IsNullOrWhiteSpace(value))                cmd.Parameters.Add(key, dbType).Value = DBNull.Value;            else                cmd.Parameters.Add(key, dbType).Value = value.Trim();        }        public struct ValuePair        {            public SqlDbType value1;            public string value2;        }




        #endregion
        #region Export DB data to Excel files        private void btnExport_Click(object sender, EventArgs e)        {            string path = @"C:\Indeed\ExportFromDB_Result";            if (!Directory.Exists(path))            {                Directory.CreateDirectory(path);            }

            SqlConnection con = new SqlConnection(@"Data Source=DEVSERVER\SQL2014;Initial Catalog=Indeed;Persist Security Info=True;User ID=sa;Password=admin12345!;");            SqlCommand cmd = new SqlCommand("Indeed_SelectAll", con);            cmd.CommandType = CommandType.StoredProcedure;            SqlDataAdapter ad = new SqlDataAdapter(cmd);            DataTable expdt = new DataTable();            ad.Fill(expdt);            if (expdt.Rows.Count > 0)            {                XLWorkbook wb = new XLWorkbook();                wb.Worksheets.Add(expdt, "Result");                wb.SaveAs(path + "\\" + "Indeed_All.xlsx");                Process.Start(path);            }        }




        #endregion
        #region Merge Excel files into one and Export        private void btnExportOneFile_Click(object sender, EventArgs e)        {            DataTable dt = new DataTable();            DataTable dtexcel = new DataTable();            DataTable dtexcelIntoOne = new DataTable();            string folderpath = @"C:\Indeed\Search_Result\2021_05_03";            string[] files = Directory.GetFiles(folderpath, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith("ls") || s.EndsWith("sx")).ToArray();            for (int s = 0; s < files.Length; s++)            {                string file = files[s].ToString();                Stream st = File.Open(file, FileMode.Open, FileAccess.Read);                IExcelDataReader reader = ExcelReaderFactory.CreateReader(st);                DataSet ds = reader.AsDataSet(new ExcelDataSetConfiguration()                {                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()                    {                        UseHeaderRow = true                    }                });                dt.Merge(ds.Tables[0]);            }            dtexcel = CreateDatatable();            for (int i = 0; i < dt.Rows.Count; i++)            {                dtexcel.Rows.Add();                dtexcel.Rows[i]["Title"] = dt.Rows[i]["Title"].ToString();                dtexcel.Rows[i]["Company"] = dt.Rows[i]["Company"].ToString();                dtexcel.Rows[i]["都道府県"] = dt.Rows[i]["都道府県"].ToString();                dtexcel.Rows[i]["Location"] = dt.Rows[i]["Location"].ToString();                dtexcel.Rows[i]["お問い合わせのURL"] = dt.Rows[i]["お問い合わせのURL"].ToString();            }            //if (dtexcel.Rows.Count > 0)            //{            //    dtexcelIntoOne = dtexcel.AsEnumerable()            //                                     .GroupBy(x => x.Field<string>("Company"))            //                                     .Select(x => x.First())            //                                     .CopyToDataTable();            //}            string filePath = @"C:\Indeed";            if (!Directory.Exists(filePath))            {                Directory.CreateDirectory(filePath);            }            SaveFileDialog savedialog = new SaveFileDialog();            savedialog.Filter = "Excel Files|*.xlsx;";            savedialog.Title = "Save";            savedialog.FileName = "Result_ExcelIntoOneFile_" + System.DateTime.Now.ToString();            savedialog.InitialDirectory = filePath;            savedialog.RestoreDirectory = true;            if (savedialog.ShowDialog() == DialogResult.OK)            {                if (Path.GetExtension(savedialog.FileName).Contains(".xlsx"))                {                    using (XLWorkbook wb = new XLWorkbook())                    {                        wb.Worksheets.Add(dtexcelIntoOne, "Result");                        wb.SaveAs(savedialog.FileName);                    }                    Process.Start(Path.GetDirectoryName(savedialog.FileName));                }            }        }


        #endregion    }}