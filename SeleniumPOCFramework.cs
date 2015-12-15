using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Interactions.Internal;
using OpenQA.Selenium.Support.UI;
using Selenium;
using Selenium.Internal;




namespace Framework
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ObjectLibrary.Helper hlpr = new ObjectLibrary.Helper();
                string batchfile = ConfigurationManager.AppSettings["batchfile"];
                string brwsr = ConfigurationManager.AppSettings["browser"];
                string testfolder = ConfigurationManager.AppSettings["testfolder"];
                string ffbin = ConfigurationManager.AppSettings["ffbin"];
                DataTable dtbatch = hlpr.dtFromExcelFile(batchfile, "BatchSheet");
                IWebDriver drv = null;
                IWebElement elem = null;
             ///   List<IWebElement> Iwebcollection = null;
                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> Iwebcollection = null;


                foreach (DataRow dr in dtbatch.Rows)
                {

                    string flagexec = dr["executeflag"].ToString();
                    if (flagexec.ToLower() == "y")
                    {
                        string scriptname = Path.Combine(testfolder, dr["scriptname"].ToString());

                        #region SCRIPTEXECUTION
                        DataTable dtscript = hlpr.dtFromExcelFile(scriptname, "Sheet1");
                        WebDriverWait wait = null;
                        foreach (DataRow drscript in dtscript.Rows)
                        {
                            if (drscript["Comment"].ToString() != null)
                            {
                                string comment = drscript["Comment"].ToString();
                            }
                            string keyword = drscript["Keyword"].ToString();
                            string url = drscript["URL"].ToString();
                            string index = drscript["Index"].ToString();
                            string fieldname = drscript["FieldName"].ToString();
                            string subcontrol = drscript["Subcontrol"].ToString();
                            string searchby = drscript["SearchBy"].ToString();
                            string searchvalue = drscript["SearchValue"].ToString();
                            string datavalue = drscript["DataValue"].ToString();
                            string testcaseid = drscript["testcaseID"].ToString();
                            string dynatext = drscript["DynaText"].ToString();


                            switch (keyword.ToLower())
                            {
                                #region LaunchBrowser
                                case "launchbrowser":
                                    {
                                        try
                                        {
                                            if (brwsr.ToLower() == "ie")
                                            {
                                                drv = new InternetExplorerDriver();
                                                wait = new WebDriverWait(drv, TimeSpan.FromMinutes(5.00));
                                            }
                                            else if (brwsr.ToLower() == "firefox")
                                            {
                                                FirefoxBinary bin = new FirefoxBinary(ffbin);
                                                FirefoxProfile ffprofile = new FirefoxProfile();
                                                drv = new FirefoxDriver(bin,ffprofile);
                                                wait = new WebDriverWait(drv, TimeSpan.FromMinutes(5.00));
                                            }
                                            else if (brwsr.ToLower() == "chrome")
                                            {

                                                drv = new ChromeDriver();
                                                wait = new WebDriverWait(drv, TimeSpan.FromMinutes(5.00));

                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            hlpr.LogtoTextFile("Exception from keyword launchbrowser: " + ex.ToString());
                                        }
                                        break;
                                    }
                                #endregion LaunchBrowser

                                #region Navigate
                                case "navigatetourl":
                                    {
                                        try
                                        {
                                            drv.Navigate().GoToUrl(url);
                                        }
                                        catch (Exception ex)
                                        {
                                            hlpr.LogtoTextFile("Exception from keyword navigatetourl: " + ex.ToString());
                                        }

                                        break;

                                    }
                                #endregion Navigate

                                #region ClickLink
                                case "clicklink":
                                    {
                                        try
                                        {
                                            hlpr.LogtoTextFile("Looking up for " + fieldname);
                                            switch (searchby.ToLower())
                                            {
                                                case "linktext":
                                                    {
                                                        try
                                                        {
                                                            elem = wait.Until(ExpectedConditions.ElementExists(By.LinkText(searchvalue)));
                                                            elem.Click();
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                                case "partiallinktext":
                                                    {
                                                        try
                                                        {
                                                            elem = wait.Until(ExpectedConditions.ElementExists(By.PartialLinkText(searchvalue)));
                                                            elem.Click();
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                                case "divtitle":
                                                    {
                                                        try
                                                        {
                                                            drv.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(40));
                                                            Iwebcollection = drv.FindElements(By.TagName("div"));
                                                            foreach (IWebElement inddivelem in Iwebcollection)
                                                            {
                                                                // inddivelem.FindElement(By.TagName(searchvalue))
                                                                //  elem = wait.Until(ExpectedConditions.ElementExists(By.TagName(searchvalue)));
                                                             //   elem = wait.Until(ExpectedConditions.ElementExists(By.TagName("div")));
                                                                if (inddivelem.GetAttribute("title") == searchvalue)
                                                                {
                                                                    inddivelem.Click();
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            hlpr.LogtoTextFile("Exception from keyword clicklink " + ex.Message);
                                        }
                                        break;
                                    }
                                #endregion Clicklink

                                #region EnterText
                                case "entertext":
                                    {
                                        try
                                        {
                                            hlpr.LogtoTextFile("Looking up for " + fieldname);
                                            switch (searchby.ToLower())
                                            {
                                                case "name":
                                                    {
                                                        try
                                                        {

                                                            elem = wait.Until(ExpectedConditions.ElementExists(By.Name(searchvalue)));
                                                            enterdata(elem, datavalue,dynatext);
                                                            

                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                                case "id":
                                                    {
                                                        try
                                                        {
                                                            elem = wait.Until(ExpectedConditions.ElementExists(By.Id(searchvalue)));
                                                            enterdata(elem, datavalue, dynatext);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                                case "xpath":
                                                    {
                                                        try
                                                        {
                                                            elem = wait.Until(ExpectedConditions.ElementExists(By.XPath(searchvalue)));
                                                            enterdata(elem, datavalue, dynatext);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            hlpr.LogtoTextFile("Exception from keyword clicklink " + ex.Message);
                                        }
                                        break;
                                    }
                                #endregion EnterText

                                #region EnterTextArea
                                case "entertextarea" :
                                    {
                                        try
                                        {
                                            hlpr.LogtoTextFile("Looking up for " + fieldname);
                                            switch (searchby.ToLower())
                                            {
                                                case "name":
                                                    {
                                                        try
                                                        {

                                                            elem = wait.Until(ExpectedConditions.ElementExists(By.Name(searchvalue)));
                                                            enterdata(elem, datavalue, dynatext);


                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                                case "id":
                                                    {
                                                        try
                                                        {
                                                            elem = wait.Until(ExpectedConditions.ElementExists(By.Id(searchvalue)));
                                                            enterdata(elem, datavalue, dynatext);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                                case "xpath":
                                                    {
                                                        try
                                                        {
                                                            elem = wait.Until(ExpectedConditions.ElementExists(By.XPath(searchvalue)));
                                                            enterdata(elem, datavalue, dynatext);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            hlpr.LogtoTextFile("Exception from keyword clicklink " + ex.Message);
                                        }
                                        break;
                                    }
                                #endregion 

                                #region ClickButton
                                case "clickbutton":
                                    {
                                        hlpr.LogtoTextFile("Looking up for " + fieldname);
                                        try
                                        {
                                            switch (searchby.ToLower())
                                            {
                                                case "name":
                                                    {
                                                        try
                                                        {
                                                            if (datavalue == "1")
                                                            {
                                                                elem = wait.Until(ExpectedConditions.ElementExists(By.Name(searchvalue)));
                                                                elem.Click();
                                                            }


                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                                case "id":
                                                    {
                                                        try
                                                        {
                                                            if (datavalue == "1")
                                                            {
                                                                elem = wait.Until(ExpectedConditions.ElementExists(By.Id(searchvalue)));
                                                                elem.Click();
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                                case "xpath":
                                                    {
                                                        try
                                                        {
                                                            if (datavalue == "1")
                                                            {
                                                                elem = wait.Until(ExpectedConditions.ElementExists(By.XPath(searchvalue)));
                                                                elem.Click();
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                                case "value":
                                                    {
                                                        try
                                                        {
                                                            drv.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(40));
                                                            Iwebcollection = drv.FindElements(By.TagName("input"));
                                                            foreach (IWebElement inddivelem in Iwebcollection)
                                                            {
                                                                // inddivelem.FindElement(By.TagName(searchvalue))
                                                                //  elem = wait.Until(ExpectedConditions.ElementExists(By.TagName(searchvalue)));
                                                                elem = wait.Until(ExpectedConditions.ElementExists(By.TagName("input")));
                                                                if (inddivelem.GetAttribute("value") == searchvalue)
                                                                {
                                                                    inddivelem.Click();
                                                                    break;

                                                                }
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            hlpr.LogtoTextFile("Unable to find " + searchby + " using " + searchvalue + "  " + ex.Message);
                                                        }
                                                        break;
                                                    }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            hlpr.LogtoTextFile("Exception from keyword clicklink " + ex.Message);
                                        }
                                        break;
                                    }
                                #endregion ClickButton

                                #region SelectRadioButton

                                #endregion SelectRadioButton

                                #region VerifyTextonPage
                                case "verifytextonpage":
                                    {

                                        string acttext = drv.FindElement(By.TagName("body")).Text;
                                        hlpr.counter = 1;
                                        hlpr.AreEqual(testcaseid, fieldname, "text", searchvalue, acttext, ObjectLibrary.Helper.CompareType.contains);
                                        break;
                                    }

                                #endregion VerifyTextPage

                                #region VerifyTable

                                #endregion VerifyTable

                                default:
                                    {

                                        hlpr.LogtoTextFile("Not a Valid keyword " + keyword);
                                        break;
                                    }

                            }



                        }
                        #endregion SCRIPTEXECUTION
                    }
                    System.Threading.Thread.Sleep(5000);
                    if (drv != null)
                    {
                        drv.Quit();
                    }
                }
                hlpr.LogtoFileCSV(hlpr.dtRep);
                Console.WriteLine("Exceution Complete Check Results at Configured Paths");

                System.Threading.Thread.Sleep(3000);
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception in Framewok " + ex.Message);
            }
        }

        static void enterdata(IWebElement elm, string val, string dynval)
        {
            if (dynval.Length > 0)
            {
                if ( dynval.ToLower() == "timestamp")
                {
                    DateTime dt = DateTime.Now;
                    elm.SendKeys(val+" " + dt.ToString("dd-MMM-yyyy hh:mm:ss"));
                }
                if (dynval.ToLower() == "guid")
                {
                    Guid g = Guid.NewGuid();
                    elm.SendKeys(val+"  "+  g.ToString().Substring(1,12));
                }
            }
            else
            {
                elm.SendKeys(val);
            }
        }
    }
}
