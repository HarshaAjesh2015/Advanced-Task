using AdvancedTask.Global;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumExtras.PageObjects;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace AdvancedTask.Pages
{
    internal class ManageListings
    {
        public ManageListings()
        {
            PageFactory.InitElements(Global.GlobalDefinitions.driver, this);
        }

        //Click on Manage Listings Link
        [FindsBy(How = How.XPath, Using = "//*[@id='account-profile-section']/div/section[1]/div/a[3]")]
        private IWebElement manageListingsLink { get; set; }

        //View the listing
        //[FindsBy(How = How.ClassName, Using = "eye icon")]
        //private IWebElement view { get; set; }

        //Delete the listing
        //[FindsBy(How = How.XPath, Using = "//*[@id='listing-management-section']/div[2]/div[1]/div[1]/table/tbody/tr[1]/td[8]/div/button[3]/i")]
        //private IWebElement delete { get; set; }

        //Edit the listing
        [FindsBy(How = How.XPath, Using = "(//i[@class='outline write icon'])[1]")]
        private IWebElement edit { get; set; }

        //Click on Yes
        [FindsBy(How = How.XPath, Using = "/html/body/div[2]/div/div[3]/button[2]")]
        private IWebElement clickYesButton { get; set; }
        //Click on No
        [FindsBy(How = How.XPath, Using = "/html/body/div[2]/div/div[3]/button[1]")]
        private IWebElement clickNoButton { get; set; }

        //Find element Chat
        [FindsBy(How = How.XPath, Using = "//*[@id='service-detail-section']/div[2]/div/div[2]/div[1]/div[2]/div/h3")]
        private IWebElement avgRatingBtn { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id='listing-management-section']/div[2]/div[1]/div[1]/table")]
        private IWebElement baseTable { get; set; }

        internal void viewListings()
        {
            //try
            //{
            GlobalDefinitions.ExcelLibrary.PopulateInCollection(Base.ExcelPath, "ManageListings");
            manageListingsLink.Click();
            Global.GlobalDefinitions.wait(3);
            var viewTitle = GlobalDefinitions.ExcelLibrary.ReadData(2, "Title");

            IList<IWebElement> allRows = baseTable.FindElements(By.TagName("tr"));
            for (int i = 1; i < allRows.Count; i++)
            {
                IList<IWebElement> cells = allRows[i].FindElements(By.TagName("td"));
                if (cells[2].Text == viewTitle)
                {
                    Thread.Sleep(2000);
                    var viewBtn = baseTable.FindElement(By.XPath($"//*[@id='listing-management-section']/div[2]/div[1]/div[1]/table/tbody/tr[{i}]/td[8]/div/button[1]"));
                    Global.GlobalDefinitions.wait(2);
                    viewBtn.Click();
                    break;
                }

            }
            Thread.Sleep(2000);
            Assert.IsTrue(avgRatingBtn.Text == "Average Ratings");

        }
        internal void deleteListings()
        {
            manageListingsLink.Click();
            //Populate the Excel Sheet
            GlobalDefinitions.ExcelLibrary.PopulateInCollection(Base.ExcelPath, "ManageListings");
            Global.GlobalDefinitions.wait(2);
            var viewTitle = GlobalDefinitions.ExcelLibrary.ReadData(2, "Title");
            IList<IWebElement> allRows = baseTable.FindElements(By.TagName("tr"));

            for (int i = 1; i < allRows.Count; i++)
            {
                IList<IWebElement> cells = allRows[i].FindElements(By.TagName("td"));
                if (cells[2].Text == viewTitle)
                {
                    Global.GlobalDefinitions.wait(2);
                    var deleteBtn = baseTable.FindElement(By.XPath($"//*[@id='listing-management-section']/div[2]/div[1]/div[1]/table/tbody/tr[{i}]/td[8]/div/button[3]/i"));
                    deleteBtn.Click();
                    Global.GlobalDefinitions.wait(2);
                    var selectAction = GlobalDefinitions.ExcelLibrary.ReadData(2, "Deleteaction");
                    Global.GlobalDefinitions.wait(2);
                    if (selectAction == "Yes")
                    {

                        clickYesButton.Click();
                    }
                    else
                    {
                        Thread.Sleep(2000);
                        clickNoButton.Click();
                    }
                    break;

                }

            }


        }

    }



}