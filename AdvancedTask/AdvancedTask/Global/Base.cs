using AdvancedTask.Configuration;
using AdvancedTask.Pages;
using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Text;
using static AdvancedTask.Global.GlobalDefinitions.ExcelLibrary;

namespace AdvancedTask.Global
{
    class Base
    {
        //To access path from Resource File

        public static string ExcelPath = Resource.ExcelPath;
        public static string ScreenshotPath = Resource.ScreenShotPath;
        public static string ReportPath = Resource.ReportPath;

        //Reports
        public static ExtentTest test;
        public static ExtentReports extent;

        //setup and tear down
        [SetUp]
        public void initialise()
        {
            GlobalDefinitions.driver = new ChromeDriver();

            GlobalDefinitions.driver.Manage().Window.Maximize();


            if (Resource.IsLogin == "true")
            {
                SignIn signinObj = new SignIn();
                signinObj.LoginSteps();

            }
            else
            {
                SignUp signupObj = new SignUp();
                signupObj.Register();


            }
        }

        [TearDown]
        public void TearDown()
        {
            // Screenshot
            String img = SaveScreenShotClass.SaveScreenshot(GlobalDefinitions.driver, "Report");//AddScreenCapture(@"E:\Dropbox\VisualStudio\Projects\Beehive\TestReports\ScreenShots\");
            test.Log(LogStatus.Info, "Image example: " + img);
            // end test. (Reports)
            extent.EndTest(test);
            // calling Flush writes everything to the log file (Reports)
            extent.Flush();
            // Close the driver :)            
            GlobalDefinitions.driver.Close();
            GlobalDefinitions.driver.Quit();
        }
    }
}
