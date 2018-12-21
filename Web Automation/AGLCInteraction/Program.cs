using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Windows.Forms;
using OpenQA.Selenium.Support.UI;
using System.IO;
using OpenQA.Selenium.Interactions;
using System.Configuration;
//Fixed the headless issue

namespace AGLCInteraction
{
    class Program
    {
        private static StreamWriter log;
        private static StreamWriter statuslog;
        private static Dictionary<string,Dictionary<string,string>> storeIdCredentails = new Dictionary<string, Dictionary<string, string>>();
        static void Main(string[] args)
        {
            IWebDriver driver = null;
            try
            {
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("******************************LTO Automation Tool********************************");
                Console.WriteLine("==================================================================================");

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Warning: Running this tool will create & submit order in the AGLC Connect Website.");
                Console.ForegroundColor = ConsoleColor.DarkGreen;

                Console.WriteLine("Do you want to run the app in background mode (Y/N)?");
                string input = Console.ReadLine();
                Console.WriteLine("==================================================================================");
                Console.WriteLine("Processing Started - Please wait for processing to be completed.");
                Console.WriteLine("==================================================================================");
                

                ChromeOptions option = new ChromeOptions();
                if(!string.IsNullOrEmpty(input))
                {
                    input = input.ToLower().Trim();
                    if (input.Equals("y"))
                    {
                        option.AddArgument("--headless");
                    }
                }

                if(!LogMessages())
                {
                    log.WriteLine("Failed to create log files.");
                    return;
                }

                if(!createStatusSheet())
                {
                    log.WriteLine("Failed to create status csv file.");
                    return;
                }

                log.WriteLine("------------------Operation Started----------------");

                log.WriteLine("Start - Credentials read.");
                if (!ReadCredentialCSVFile())
                {
                    log.WriteLine("Failed Reading the Credentials CSV file.");
                    return;
                }
                log.WriteLine("End - Credentials read.");

                string directoryPath = DateTime.Now.ToString("yyyy/MM/dd_hh:mm:ss").Replace("/", "");
                directoryPath = "./Working/Processed/" + directoryPath.Replace(":", "");
                DirectoryInfo di = Directory.CreateDirectory(directoryPath);
                //Excel Header
                statuslog.WriteLine("Store Number," + "Incoming Case Count," + "Incoming Total," + "Process Start," + "Process End," + "Process Status");
                
                foreach (var store in storeIdCredentails)
                {
                    driver = new ChromeDriver(option);
                    string filePath = "./Working/" + store.Key + ".csv";
                    if(!File.Exists(filePath))
                    {
                        log.WriteLine("File does not exists - {0}", filePath);
                        var logLine = store.Key + ",,," + DateTime.Now.ToString("yyyy/MM/dd_hh:mm") + "," + DateTime.Now.ToString("yyyy/MM/dd_hh:mm") + ","+ "Failed";
                        statuslog.WriteLine(logLine);
                        continue;
                    }
                    filePath = Path.GetFullPath(filePath);
                    var creds = store.Value;
                    string startTime = "";
                    string Qty ="", TotalCost = "";
                    foreach (var cred in creds)//Assuming always one
                    {
                        if(cred.Key != null && cred.Value != null)
                        {
                            try
                            {
                                log.WriteLine("Processing Started for Store {0}", store.Key);
                                startTime = DateTime.Now.ToString("yyyy/MM/dd_hh:mm");
                                createDriveExecuteTest(filePath, cred.Key, cred.Value, driver, out Qty, out TotalCost);
                                var logLine = store.Key + "," + Qty + "," + TotalCost + "," + startTime + "," + DateTime.Now.ToString("yyyy/MM/dd_hh:mm") + "," + "Passed";
                                statuslog.WriteLine(logLine);
                                log.WriteLine("Processing Ended for Store {0}", store.Key);
                                System.IO.File.Move(filePath, directoryPath + "//" + store.Key + ".csv");
                            }
                            catch( Exception ex)
                            {
                                log.WriteLine("Processing failed for the store {0}", store.Key);
                                log.WriteLine(ex.Message + ex.InnerException);
                                var logLine = store.Key + "," + Qty + "," + TotalCost + "," + startTime + "," + DateTime.Now.ToString("yyyy/MM/dd_hh:mm") + "," + "Failed";
                                statuslog.WriteLine(logLine);
                            }
                            finally
                            {
                                if (driver != null)
                                {
                                    driver.Close();
                                }
                            }
                        }
                        else
                        {
                            log.WriteLine("Invalid Credentials for store {0}", store.Key);
                        }
                    }
                }
                log.WriteLine("------------------Operation Completed----------------");
                Console.WriteLine("==================================================================================");
                Console.WriteLine("Processing Completed - Please check the status file.");
                Console.WriteLine("==================================================================================");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                if (driver != null)
                {
                    driver.Quit();
                }
                log.WriteLine(ex.Message + ex.InnerException);
            }
            finally
            {
                if (driver != null)
                {
                    driver.Quit();
                }
                log.Close();
                statuslog.Close();
            }
        }

        private static void createDriveExecuteTest(string filePath, string userName, string pwd, IWebDriver driver, out string Qty, out string TotalCost)
        {
            string url = ConfigurationManager.AppSettings["URL"];
            driver.Navigate().GoToUrl(url);

            if (driver.FindElement(By.Id("ctl00_IdWelcome_ExplicitLogin")).Displayed)
            {
                driver.FindElement(By.Id("ctl00_IdWelcome_ExplicitLogin")).Click();
            }

            if (driver.FindElement(By.Id("ctl00_PlaceHolderMain_signInControl_UserName")).Displayed)
            {
                var UserNameBox = driver.FindElement(By.Id("ctl00_PlaceHolderMain_signInControl_UserName"));
                UserNameBox.SendKeys(userName);
            }

            if (driver.FindElement(By.Id("ctl00_PlaceHolderMain_signInControl_Password")).Displayed)
            {
                var pwdBox = driver.FindElement(By.Id("ctl00_PlaceHolderMain_signInControl_Password"));
                pwdBox.SendKeys(pwd);
            }


            driver.FindElement(By.Id("ctl00_PlaceHolderMain_signInControl_Login")).Click();
            System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
            driver.FindElement(By.XPath("//*[@id='topNav']/ul[1]/li[5]/a[1]")).Click();

            bool isUpdate = false;

            var addOnButtons = driver.FindElements(By.CssSelector("input[id *= _AddToOrderImageButton]"));
            foreach(var btn in addOnButtons)
            {
                if (btn.Displayed && btn.Enabled)
                {
                    log.WriteLine("Add on present for the store file {0}", filePath);
                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                    js.ExecuteScript("arguments[0].scrollIntoView(true);", btn);
                    System.Threading.Thread.Sleep(3000);
                    btn.Click();
                    isUpdate = true;
                    break;
                }
            }


            if(isUpdate == false)
            {
                driver.FindElement(By.Id("ctl00_ctl45_g_c0926b90_c918_42dc_bd19_e0920d569bfc_ctl00_ImportOrderHyperLink")).Click();
                System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
                var modal = driver.FindElement(By.Id("modalWindow"));
                //To find the iframe
                IWebElement Object = driver.FindElement(By.Id("jqmContent"));
                //To switch to and set focus to the iframe
                driver.SwitchTo().Frame(Object);
                IWebElement startb = driver.FindElement(By.Id("ctl01"));
                //startb.FindElement(By.Id("ProductFileUpload")).Click();
                var fileupload = startb.FindElement(By.Id("ProductFileUpload"));
                fileupload.SendKeys(filePath);


                System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
                //SendKeys.SendWait(filePath);
                SendKeys.SendWait(@"{Enter}");
                startb.FindElement(By.Id("UploadButton")).Click();

                System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
                IWebElement CreateNewOrderButton = driver.FindElement(By.Id("CreateNewOrderButton"));
                CreateNewOrderButton.Click();
                System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(10).TotalMilliseconds);

                //var AdditionalAddresses = driver.FindElement(By.Id("ctl00_ctl45_g_8929e8c5_f33b_4bf7_8f54_f558a5e1b1cf_ctl00_AdditionalAddressesTextBox"));
                //AdditionalAddresses.SendKeys("Regina.Tsai@lsgp.ca");
                //driver.FindElement(By.Id("ctl00_ctl45_g_8929e8c5_f33b_4bf7_8f54_f558a5e1b1cf_ctl00_SaveDetailsLinkButton")).Click();
                //driver.FindElement(By.XPath("//*[@id='ctl00_ctl45_g_8929e8c5_f33b_4bf7_8f54_f558a5e1b1cf_ctl00_BackToOrdersHyperLink']/span[2]")).Click();
                //driver.FindElement(By.Id("zz4_Menu_t")).Click();
                //System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(3).TotalMilliseconds);
                ////driver.FindElement(By.XPath("//*[@id='zz4_Menu_t/']")).Click();
                //System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(3).TotalMilliseconds);
            }
            else
            {
                driver.FindElement(By.Id("ctl00_ctl45_g_8929e8c5_f33b_4bf7_8f54_f558a5e1b1cf_ctl00_ImportOrderHyperLink")).Click();
                System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
                var modal = driver.FindElement(By.Id("modalWindow"));
                //To find the iframe
                IWebElement Object = driver.FindElement(By.Id("jqmContent"));
                //To switch to and set focus to the iframe
                driver.SwitchTo().Frame(Object);
                IWebElement startb = driver.FindElement(By.Id("ctl01"));
                //startb.FindElement(By.Id("ProductFileUpload")).Click();
                var fileUpload = startb.FindElement(By.Id("ProductFileUpload"));
                fileUpload.SendKeys(filePath);

                System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
                //SendKeys.SendWait(filePath);
                SendKeys.SendWait(@"{Enter}");
                startb.FindElement(By.Id("UploadButton")).Click();

                System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
                IWebElement AddToOrderButton = driver.FindElement(By.Id("AddToOrderButton"));
                AddToOrderButton.Click();
                System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(10).TotalMilliseconds);
            }


            Qty = driver.FindElement(By.Id("ctl00_ctl45_g_8929e8c5_f33b_4bf7_8f54_f558a5e1b1cf_ctl00_EstimatedTotalQuantityTopLabel")).GetAttribute("textContent");
            Qty = Qty.Replace(",", "");
            TotalCost = driver.FindElement(By.Id("ctl00_ctl45_g_8929e8c5_f33b_4bf7_8f54_f558a5e1b1cf_ctl00_EstimatedTotalPriceTopLabel")).GetAttribute("textContent");
            TotalCost = TotalCost.Replace(",", "");
            System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
            //submit button
            bool isSubmit = Convert.ToBoolean(ConfigurationManager.AppSettings["SubmitOrder"]);
            if(isSubmit)
            {
                driver.FindElement(By.XPath("//*[@id='ctl00_ctl45_g_8929e8c5_f33b_4bf7_8f54_f558a5e1b1cf_ctl00_SubmitOrderLinkButton']/span[2]")).Click();
                IAlert alert = driver.SwitchTo().Alert();
                alert.Accept();
                System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
            }
            System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(5).TotalMilliseconds);
            driver.FindElement(By.Id("zz4_Menu_t")).Click();
            System.Threading.Thread.Sleep((int)System.TimeSpan.FromSeconds(2).TotalMilliseconds);
            IWebElement personalMenu = driver.FindElement(By.Id("zz4_Menu_t"));
            personalMenu.FindElement(By.Id("zz3_ID_Logout")).Click();
        }

        private static bool ReadCredentialCSVFile()
        {
            bool result;
            try
            {
                using (var fs = new FileStream("./Store/Credentials.csv", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var rd = new StreamReader(fs, Encoding.Default))
                {
                    while (!rd.EndOfStream)
                    {
                        try
                        {
                            var splits = rd.ReadLine().Split(',');
                            var cred = new Dictionary<string, string>();
                            cred.Add(splits[1], splits[2]);
                            storeIdCredentails.Add(splits[0], cred);
                        }
                        catch(Exception ex)
                        {
                            log.WriteLine("Exception occured while reading the Credentials file.", ex.ToString());
                        }
                    }
                }
                result = true;
            }
            catch(Exception ex)
            {
                log.WriteLine("Exception occured while reading the Credentials file.", ex.ToString());
                result = false;
            }
            return result;
        }

        private static bool LogMessages()
        {
            try
            {
                var fileName = DateTime.Now.ToString().Replace("/", "");
                fileName = fileName.Replace(" ", "");
                fileName = fileName.Replace(":", "");
                var logFilePath = "./Store/Logs/" + fileName + ".txt";
                logFilePath = Path.GetFullPath(logFilePath);
                log = new StreamWriter(logFilePath);
                log.WriteLine(DateTime.Now);
                log.WriteLine("--------------------");
                return true;
            }
            catch(Exception ex)
            {
                log.WriteLine("Exception occured while LogMessages", ex.ToString());
                return false;
            }
        }

        private static bool createStatusSheet()
        {
            try
            {
                var fileName = DateTime.Now.ToString("yyyy/MM/dd hh:mm").Replace("/", "");
                fileName = fileName.Replace(" ", "_");
                fileName = fileName.Replace(":", "");
                fileName = "LTO_OrderSummary_" + fileName;
                var logFilePath = "./Store/Status" + fileName + ".csv";
                logFilePath = Path.GetFullPath(logFilePath);
                statuslog = new StreamWriter(logFilePath);
                return true;
            }
            catch (Exception ex)
            {
                log.WriteLine("Exception occured while LogMessages", ex.ToString());
                return false;
            }
        }
    }
}
