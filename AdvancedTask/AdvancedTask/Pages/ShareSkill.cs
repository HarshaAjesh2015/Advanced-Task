using AdvancedTask.Global;
using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

using SeleniumExtras.PageObjects;

namespace AdvancedTask.Pages
{
    internal class ShareSkill
    {
        public ShareSkill()
        {
            PageFactory.InitElements(Global.GlobalDefinitions.driver, this);
        }

        //Click on ShareSkill Button
        
        [FindsBy(How = How.XPath, Using = "//*[@id='account-profile-section']/div/section[1]/div/div[2]/a")]
       private IWebElement ShareSkillButton { get; set; }

        //Enter the Title in textbox
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[1]/div/div[2]/div/div[1]/input")]
        private IWebElement Title { get; set; }

        //Enter the Description in textbox
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[2]/div/div[2]/div[1]/textarea")]
        private IWebElement Description { get; set; }

        //Click on Category Dropdown
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[3]/div[2]/div/div/select")]
        private IWebElement CategoryDropDown { get; set; }

        //Click on SubCategory Dropdown
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[3]/div[2]/div/div[2]/div[1]/select")]
        private IWebElement SubCategoryDropDown { get; set; }

        //Enter Tag names in textbox
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[4]/div[2]/div/div/div/div/input")]
        private IWebElement Tags { get; set; }

        ////Select the Service type Hourly basis
        //[FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section]/div[2]/div/form/div[5]/div[2]/div[1]/div[1]/div/input")]
        //private IWebElement ServiceTypeHourly { get; set; }

        ////Select the Service type One-off
        //[FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section]/div[2]/div/form/div[5]/div[2]/div[1]/div[2]/div/input")]
        //private IWebElement ServiceTypeOneoff { get; set; }

        //Select the Service Type
        [FindsBy(How = How.Name, Using = "serviceType")]
        private IList<IWebElement> serviceOptions { get; set; }

        //Select the Location Type
        //[FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[6]/div[2]/div/div[2]/div/input")]
        //private IWebElement LocationTypeOption { get; set; }

        [FindsBy(How = How.Name, Using = "locationType")]
        private IList<IWebElement> LocationTypeOption { get; set; }

        //Click on Start Date dropdown
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[7]/div[2]/div/div[1]/div[2]/input")]
        private IWebElement StartDateDropDown { get; set; }

        //Click on End Date dropdown
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[7]/div[2]/div/div[1]/div[4]/input")]
        private IWebElement EndDateDropDown { get; set; }

        //Storing the table of available days
        [FindsBy(How = How.Name, Using = "Available")]
        private IList<IWebElement> Days { get; set; }

        //Storing the starttime
        [FindsBy(How = How.Name, Using = "StartTime")]
        private IList<IWebElement> StartTime { get; set; }

        [FindsBy(How = How.Name, Using = "EndTime")]
        private IList<IWebElement> EndTime { get; set; }

        //Click on StartTime dropdown
        //[FindsBy(How = How.Name, Using = "StartTime")]
        //private IWebElement StartTimeDropDown { get; set; }

        ////Click on EndTime dropdown
        //[FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[7]/div[2]/div/div[2]/div[3]/input")]
        //private IWebElement EndTimeDropDown { get; set; }

        //Click on Skill Trade option
        [FindsBy(How = How.Name, Using = "skillTrades")]
        private IList<IWebElement> SkillTradeOption { get; set; }

        //Enter Skill Exchange
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[8]/div[4]/div/div/div/div/div/input")]
        private IWebElement SkillExchange { get; set; }

        //Enter the amount for Credit
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[8]/div[4]/div/div/input")]
        private IWebElement CreditAmount { get; set; }

        //Click on Active/Hidden option
        [FindsBy(How = How.Name, Using = "isActive")]
        private IList<IWebElement> ActiveOption { get; set; }

        //Click on Save button
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[11]/div/input[1]")]
        private IWebElement Save { get; set; }

        //Click on Cancel button
        [FindsBy(How = How.XPath, Using = "//*[@id='service-listing-section']/div[2]/div/form/div[11]/div/input[2]")]
        private IWebElement Cancel { get; set; }

        //Find Manage Listings Page
        [FindsBy(How = How.XPath, Using = "//*[@id='listing-management-section']/div[2]/h2")]
        private IWebElement manageListingsText { get; set; }

        internal void EnterShareSkill()
        {
            try
            {

                ShareSkillButton.Click();
                GlobalDefinitions.ExcelLibrary.PopulateInCollection(Base.ExcelPath, "ShareSkill");
                Global.GlobalDefinitions.wait(2);
                Title.SendKeys(GlobalDefinitions.ExcelLibrary.ReadData(2, "Title"));
                Description.SendKeys(GlobalDefinitions.ExcelLibrary.ReadData(2, "Description"));
                CategoryDropDown.SendKeys(GlobalDefinitions.ExcelLibrary.ReadData(2, "Category"));
                SubCategoryDropDown.SendKeys(GlobalDefinitions.ExcelLibrary.ReadData(2, "SubCategory"));
                Tags.SendKeys(GlobalDefinitions.ExcelLibrary.ReadData(2, "Tags"));
                Tags.SendKeys(Keys.Enter);

                var serviceOption = GlobalDefinitions.ExcelLibrary.ReadData(2, "ServiceType");

                if (serviceOption == "Hourly basis service")
                {
                    serviceOptions[0].Click();
                }
                else
                {
                    serviceOptions[1].Click();
                }

                var locationTypeValue = GlobalDefinitions.ExcelLibrary.ReadData(2, "LocationType");

                if (locationTypeValue == "On-site")
                {
                    LocationTypeOption[0].Click();
                }
                else
                {
                    LocationTypeOption[1].Click();
                }

                StartDateDropDown.SendKeys(GlobalDefinitions.ExcelLibrary.ReadData(2, "Startdate"));
                EndDateDropDown.SendKeys(GlobalDefinitions.ExcelLibrary.ReadData(2, "Enddate"));

                //Days.SendKeys(GlobalDefinitions.ExcelLib.ReadData(2, "Selectday"));

                var availableDays = GlobalDefinitions.ExcelLibrary.ReadData(2, "Selectday");
                var availDays = availableDays.Split(new char[] { ',' });
                var startTime = GlobalDefinitions.ExcelLibrary.ReadData(2, "Starttime");
                var timeStart = DateTime.Parse(startTime).ToString("hhmmtt");
                var endTime = GlobalDefinitions.ExcelLibrary.ReadData(2, "Endtime");
                var timeEnd = DateTime.Parse(endTime).ToString("hhmmtt");

                foreach (var day in availDays)
                {
                    switch (day)
                    {
                        case "Sun":
                            Days[0].Click();
                            StartTime[0].Click();
                            StartTime[0].SendKeys(timeStart);
                            EndTime[0].Click();
                            EndTime[0].SendKeys(endTime);
                            break;
                        case "Mon":
                            Days[1].Click();
                            StartTime[1].Click();
                            StartTime[1].SendKeys(timeStart);
                            EndTime[1].Click();
                            EndTime[1].SendKeys(timeEnd);
                            break;
                        case "Tue":
                            Days[2].Click();
                            StartTime[2].Click();
                            StartTime[2].SendKeys(timeStart);
                            EndTime[2].Click();
                            EndTime[2].SendKeys(timeEnd);
                            break;
                        case "Wed":
                            Days[3].Click();
                            StartTime[3].Click();
                            StartTime[3].SendKeys(timeStart);
                            EndTime[3].Click();
                            EndTime[3].SendKeys(timeEnd);
                            break;
                        case "Thu":
                            Days[4].Click();
                            StartTime[4].Click();
                            StartTime[4].SendKeys(timeStart);
                            EndTime[4].Click();
                            EndTime[4].SendKeys(timeEnd);
                            break;
                        case "Fri":
                            Days[5].Click();
                            StartTime[5].Click();
                            StartTime[5].SendKeys(timeStart);
                            EndTime[5].Click();
                            EndTime[5].SendKeys(timeEnd);
                            break;
                        case "Sat":
                            Days[6].Click();
                            StartTime[6].Click();
                            StartTime[6].SendKeys(timeStart);
                            EndTime[6].Click();
                            EndTime[6].SendKeys(timeEnd);
                            break;
                    }


                }

                Global.GlobalDefinitions.wait(2);

                var tradeSkillOptions = GlobalDefinitions.ExcelLibrary.ReadData(2, "SkillTrade");

                if (tradeSkillOptions == "Skill-Exchange")
                {
                    SkillTradeOption[0].Click();

                    SkillExchange.SendKeys(GlobalDefinitions.ExcelLibrary.ReadData(2, "Skill-Exchange"));

                    SkillExchange.SendKeys(Keys.Enter);

                }
                else
                {
                    SkillTradeOption[1].Click();

                    CreditAmount.SendKeys(GlobalDefinitions.ExcelLibrary.ReadData(2, "Credit"));
                }

                Global.GlobalDefinitions.wait(2);

                var actOrHide = GlobalDefinitions.ExcelLibrary.ReadData(2, "Active");

                if (actOrHide == "Active")
                {
                    ActiveOption[0].Click();
                }
                else
                {
                    ActiveOption[1].Click();
                }

                var saveORcancel = GlobalDefinitions.ExcelLibrary.ReadData(2, "Save Or Cancel");

                if (saveORcancel == "Save")
                {
                    Save.Click();
                }
                else
                {
                    Cancel.Click();
                }

                Thread.Sleep(2000);

                Assert.IsTrue(manageListingsText.Text == "Manage Listings");

            }
            catch (Exception ex)
            {

                Assert.IsTrue(false);
            }
        }

        internal void EditShareSkill()
        {


        }

    }
}

    

