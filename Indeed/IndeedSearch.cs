﻿using System;



        #region Search by Key and Location on Indeed and Export to Excel 
            int start = 0;//start page = 0,next page = 10,next page = 20.......

            IWebDriver driver = new ChromeDriver();
                driver.Navigate().GoToUrl("https://jp.indeed.com/jobs?q=" + txtKeyword.Text + "&l=" + txtLocation.Text + "&limit=50&radius=25&start=" + start.ToString());//searh url with start param
               // driver.Navigate().GoToUrl("https://jp.indeed.com/%E6%B1%82%E4%BA%BA?q=%E3%83%97%E3%83%AD%E3%82%B0%E3%83%A9%E3%83%9E&l=%E5%8C%97%E6%B5%B7%E9%81%93&limit=50&radius=25");//searh url with start param

                Thread.Sleep(3000);//wait page load

                //IList<IWebElement> popup = driver.FindElements(By.ClassName("popover"));
                //if (popup.Count() > 0)
                //{
                //    Actions action = new Actions(driver);
                //    action.SendKeys(OpenQA.Selenium.Keys.Escape);
                //}

                //get all title,company,location by array
                IList<IWebElement> arrTitle = driver.FindElements(By.ClassName("title"));//募集内容 
                IList<IWebElement> arrCompany = driver.FindElements(By.ClassName("company"));//会社名
                                                                                             //IList<IWebElement> arrPrefecture = driver.FindElements(By.ClassName("prefecture"));//都道府県
                IList<IWebElement> arrLocation = driver.FindElements(By.ClassName("location"));//所在地

                for (int i = 0; i < arrTitle.Count; i++)

                
                var CountLi = 0;
                IList<IWebElement> arrPaging ;
                try
                {
                    arrPaging= driver.FindElements(By.ClassName("pn"));
                }
                catch {
                    arrPaging = driver.FindElements(By.ClassName("np"));
                }
                CountLi = arrPaging.Count()+1;

                var result = driver.FindElement(By.XPath("//*[@id=\"resultsCol\"]/nav/div/ul/li["+ CountLi + "]"));
                ////*[@id="resultsCol"]/nav/div/ul/li[1]
               // IWebElement a = driver.FindElement(By.CssSelector("ul.pagination-list li:last-child"));
                {
                //IWebElement name = a.FindElement(By.TagName("b"));
                go = false;

                //check next page exists
                if (arrPaging.Count == 1)
            if (!Directory.Exists(saveFolder))

            //using (XLWorkbook wb = new XLWorkbook())
            //{
            //    wb.Worksheets.Add(dtResult, "Indeed Search Result");
            //    wb.SaveAs(txtFilePath.Text);

            //    Process.Start(txtFilePath.Text);
            //}
        }




        #endregion
        #region Import Excel Files to DB




        #endregion
        #region Export DB data to Excel files

            SqlConnection con = new SqlConnection(@"Data Source=DEVSERVER\SQL2014;Initial Catalog=Indeed;Persist Security Info=True;User ID=sa;Password=admin12345!;");




        #endregion
        #region Merge Excel files into one and Export


        #endregion