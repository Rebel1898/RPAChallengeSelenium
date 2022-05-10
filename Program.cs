using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Data;

namespace SeleniumRPAChallenge
{
   class Program
    {
        static void Main(string[] args)
        {
            string[] arraynames = { "labelPhone", "labelAddress", "labelCompanyName", "labelFirstName", "labelEmail", "labelRole", "labelLastName" };
            string[] colnames = { "Phone Number", "Address", "Company Name", "First Name", "Email", "Role in Company", "Last Name " };
            DataTable Casos = GetDataFromExcel("./challenge.xlsx", "Sheet1");
            bool start = true;
            FirefoxDriver driver = new FirefoxDriver(@"C:\Users\X\Downloads\");
            WebDriverWait wait30 = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            wait30.IgnoreExceptionTypes(typeof(StaleElementReferenceException));
            wait30.IgnoreExceptionTypes(typeof(NoSuchElementException));
            int contador = 0;
            foreach (DataRow fila in Casos.Rows)
            {
                if (start)
                {
                    driver.Navigate().GoToUrl("https://www.rpachallenge.com/");
                    start = false;
                    driver.Manage().Window.Maximize();

                    wait30.Until(ExpectedConditions.ElementExists(By.TagName("BUTTON")));
                    var botonStart = driver.FindElement(By.TagName("BUTTON"));
                    botonStart.Click();
                }           
                var input = driver.FindElements(By.TagName("input"));
                for (int i = 0; i < input.Count; i++)
                {
                    var NAME = input[i].GetAttribute("ng-reflect-name");
                    for (int j = 0; j < arraynames.Length; j++)
                    {
                        if (NAME == arraynames[j])
                        {
                            input[i].SendKeys(fila[colnames[j]].ToString());
                            break;
                        }
                    }
                }
                if (contador < 10) { 
                wait30.Until(ExpectedConditions.ElementExists(By.CssSelector(".btn.uiColorButton")));
                driver.FindElement(By.CssSelector(".btn.uiColorButton")).Click();
                }
                contador++;
            }
            DataTable GetDataFromExcel(string path, dynamic worksheet)
            {
                DataTable dt = new DataTable();
                using (XLWorkbook workBook = new XLWorkbook(path))
                {
                    IXLWorksheet workSheet = workBook.Worksheet(worksheet);
                    bool firstRow = true;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                if (!string.IsNullOrEmpty(cell.Value.ToString()))
                                {
                                    dt.Columns.Add(cell.Value.ToString());
                                }
                                else
                                {
                                    break;
                                }
                            }
                            firstRow = false;
                        }
                        else
                        {
                            int i = 0;
                            DataRow toInsert = dt.NewRow();
                            foreach (IXLCell cell in row.Cells(1, dt.Columns.Count))
                            {
                                try
                                {
                                    toInsert[i] = cell.Value.ToString();
                                }
                                catch
                                {

                                }
                                i++;
                            }
                            dt.Rows.Add(toInsert);
                        }
                    }
                    return dt;
                }
            }
        }
    }
}

