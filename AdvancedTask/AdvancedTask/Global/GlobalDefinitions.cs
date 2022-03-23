using Excel;
using ExcelDataReader;
using Microsoft.Xrm.Sdk;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using static AdvancedTask.Global.GlobalDefinitions.ExcelLibrary;

namespace AdvancedTask.Global
{
    public class GlobalDefinitions
    {
        public static IWebDriver driver { get; set; }

        public static void wait(int time)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(time);
        }

        public static IWebElement WaitForElement(IWebDriver driver, By by, int timeinSeconds)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeinSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(by));
        }
        public class ExcelLibrary
        {
            static List<DataCollection> dataCollections = new List<DataCollection>();

            public class DataCollection
            {
                public int rowNumber { get; set; }
                public string colName{ get; set; }
                public string colValue { get; set; }
            }
            public static void ClearData()
            {
                dataCollections.Clear();
            }
            private static DataTable ExcelToDataTable(string fileName, string SheetName)
            {
                using (System.IO.FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {
                        excelReader.IsFirstRowAsColumnNames= true;

                        //Return as dataset
                        DataSet result = excelReader.AsDataSet();
                        //Get all the tables
                        DataTableCollection table = result.Tables;

                        // store it in data table
                        DataTable resultTable = table[SheetName];

                        //excelReader.Dispose();
                        //excelReader.Close();
                        // return
                        return resultTable;
                    }
                }

            }
            public static string ReadData(int rowNumber, string columnName)
            {
                try
                {
                    //Retriving Data using LINQ to reduce much of iterations

                    rowNumber = rowNumber - 1;
                    string data = (from colData in dataCollections
                                   where colData.colName == columnName && colData.rowNumber == rowNumber
                                   select colData.colValue).SingleOrDefault();

                    // var datas = dataCollections.Where(x => x.colName == columnName && x.rowNumber == rowNumber).SingleOrDefault().colValue;


                    return data.ToString();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception occurred in ExcelLib Class ReadData Method!" + Environment.NewLine + ex.Message.ToString());
                    return null;
                }

            }

            public static void PopulateInCollection(string fileName, string sheetName)
            {
                ExcelLibrary.ClearData();
                DataTable table = ExcelToDataTable(fileName, sheetName);

                //Iterate through the rows and columns of the Table
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        DataCollection dtTable = new DataCollection()
                        {
                            rowNumber = row,
                            colName = table.Columns[col].ColumnName,
                            colValue = table.Rows[row - 1][col].ToString()
                        };


                        //Add all the details for each row
                        dataCollections.Add(dtTable);

                    }
                }

            }
            public class SaveScreenShotClass
            {
                public static string SaveScreenshot(IWebDriver driver, string ScreenShotFileName) // Definition
                {
                    var folderLocation = (Base.ScreenshotPath);

                    if (!System.IO.Directory.Exists(folderLocation))
                    {
                        System.IO.Directory.CreateDirectory(folderLocation);
                    }

                    var screenShot = ((ITakesScreenshot)driver).GetScreenshot();
                    var fileName = new StringBuilder(folderLocation);

                    fileName.Append(ScreenShotFileName);
                    fileName.Append(DateTime.Now.ToString("_dd-mm-yyyy_mss"));
                    //fileName.Append(DateTime.Now.ToString("dd-mm-yyyym_ss"));
                    fileName.Append(".jpeg");
                    screenShot.SaveAsFile(fileName.ToString(), ScreenshotImageFormat.Jpeg);
                    return fileName.ToString();
                }
            }
        }
    }

}




