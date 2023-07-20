using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;
using System.Threading;

namespace Nazia_Accelerate
{
    public class Tests
    {
        IWebDriver driver;
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver("C: \\Users\\rohitkanti.das\\Downloads\\chromedriver_win32\\chromedriver.exe");
            driver.Manage().Window.Maximize();
        }

        [Test]
        public void Test1()
        {
            driver.Navigate().GoToUrl("https://amazon.in");

            var excelApp = new Application()
            {
                Visible = true
            };

            var workbook = excelApp.Workbooks.Add();
            var sheet = (Worksheet)workbook.Worksheets[1];
            sheet.Name = "Today's deals";

            string excelfilepath = "C:\\Users\\rohitkanti.das\\source\\repos\\Nazia_Accelerate\\top_deals.xlsx";
            if (File.Exists(excelfilepath))
            {
                File.Delete(excelfilepath);
            }
            workbook.SaveAs(excelfilepath);
            Thread.Sleep(5000);
            workbook.Save();
            workbook.Close();
            excelApp.Quit();
            System.Console.WriteLine("This is new Line");

            Assert.Pass();
        }
    }
}