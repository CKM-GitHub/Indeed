using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Indeed
{
    public partial class frmBaseConnect : Form
    {
        public frmBaseConnect()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            IWebDriver driver = new ChromeDriver();
            try
            {
                this.Cursor = Cursors.WaitCursor;
                DataTable dtFinal = new DataTable();
                DataTable dtResult = new DataTable();
                dtResult.Columns.Add("会社名Eng");
                dtResult.Columns.Add("会社名");
                dtResult.Columns.Add("設立年月");
                dtResult.Columns.Add("資本金");
                dtResult.Columns.Add("売上高");
                dtResult.Columns.Add("従業員数");

                ArrayList UrlList = new ArrayList();
                bool LastPage = false;
                int pgNo = 1;
                do
                {
                    driver.Navigate().GoToUrl("https://baseconnect.in/companies/category/377d61f9-f6d3-4474-a6aa-4f14e3fd9b17?page=" + pgNo.ToString());
                    Thread.Sleep(3000);//wait page load

                    IList<IWebElement> arrURL = driver.FindElements(By.ClassName("searches__result__list__header__title"));
                    foreach(IWebElement ele in arrURL)
                    {
                        IWebElement a1 = ele.FindElement(By.TagName("a"));
                        UrlList.Add(a1.GetAttribute("href"));
                    }

                    IList<IWebElement> nextpg = driver.FindElements(By.ClassName("next_page"));
                    if (nextpg.Count == 0)
                        LastPage = true;

                    //LastPage = true;
                    
                    pgNo++;
                } while (!LastPage);

                string CompanyNameEng = string.Empty;
                string CompanyName = string.Empty;
                string EstablishDate = string.Empty;
                string Capital = string.Empty;
                string SalesAmount = string.Empty;
                string Employee = string.Empty;

                int j = 0;
                foreach(string url in UrlList)
                {
                    driver.Navigate().GoToUrl(url);
                    Thread.Sleep(5000);


                    IList<IWebElement> ele1 = driver.FindElements(By.ClassName("modalBanner__show"));
                    if (ele1.Count > 0)
                    {
                        IList<IWebElement> e1 = driver.FindElements(By.ClassName("modalClose-btn"));
                        if (e1.Count > 1)
                        {
                            e1[1].Click();
                        }
                    }

                    if(driver.FindElements(By.ClassName("node__header__title__english")).Count>0)
                        CompanyNameEng = driver.FindElement(By.ClassName("node__header__title__english")).Text;
                    CompanyName = driver.FindElement(By.ClassName("node__header__text__title")).Text;
                    IList<IWebElement> we = driver.FindElements(By.ClassName("mincho"));


                    EstablishDate = string.Empty;
                    Capital = string.Empty;
                    SalesAmount = string.Empty;
                    Employee = string.Empty;

                    IWebElement div = driver.FindElement(By.ClassName("nodeTable--simple"));
                    IList<IWebElement> dl = div.FindElements(By.TagName("dl"));
                    foreach(IWebElement e1 in dl)
                    {
                        string str1 = e1.FindElement(By.TagName("dt")).Text;
                        switch (str1)
                        {
                            case "設立年月":
                                EstablishDate = e1.FindElement(By.TagName("dd")).Text;
                                break;
                            case "資本金":
                                Capital = e1.FindElement(By.TagName("dd")).Text;
                                break;
                            case "売上高":
                                SalesAmount = e1.FindElement(By.TagName("dd")).Text;
                                break;
                            case "従業員数":
                                Employee = e1.FindElement(By.TagName("dd")).Text;
                                break;
                        }
                    }

                    //EstablishDate = we[0].Text;
                    //Capital = we[2].Text;
                    //SalesAmount = we[3].Text;
                    //Employee = we[5].Text;

                    dtResult.Rows.Add();
                    dtResult.Rows[j]["会社名Eng"] = CompanyNameEng;
                    dtResult.Rows[j]["会社名"] = CompanyName;
                    dtResult.Rows[j]["設立年月"] = EstablishDate;
                    dtResult.Rows[j]["資本金"] = Capital;
                    dtResult.Rows[j]["売上高"] = SalesAmount;
                    dtResult.Rows[j]["従業員数"] = Employee;

                    j++;
                }


                string saveFolder = @"C:\Indeed\Search_Result\BC";
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
                            wb.Worksheets.Add(dtResult, "SearchResult");
                            wb.SaveAs(savedialog.FileName);
                        }
                        Process.Start(Path.GetDirectoryName(savedialog.FileName));
                    }
                }
                this.Cursor = Cursor.Current;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
            finally
            {
                driver.Quit();
            }
        }
    }
}
