using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Data;
using System.IO;
using System.Globalization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Internal;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions.Internal;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Internal;


using System.Text.RegularExpressions;
//using WatiN.Core;
using AutoItX3Lib;

#region comments
/*Version 0.0.0.1
 * Created on Feb 26, 2013
 * Author : Sneha Dhole
 * Developed from scratch
 */
/*Version 0.0.0.2
 * Created on Mar 01, 2013
 * Author: Sneha Dhole
 * Added code to search by regular expression
 * Added code for index
 * Made a common method GetWebElement for all controltypes
 * Used index and regexp as optional parameters for perfaction method
 */
/*Version 0.0.0.3
 * Created on March 11, 2013
 * Author: Sneha Dhole
 * Added function goToFirstFrame
 * Added control types img, textarea, frame, iframe and tab
 * Added src and onclick cases in searchBy conditions.
 */
/*Version 0.0.0.4
 * Created on March 22, 2013
 * Author: Sneha Dhole
 * Added function getdata
 */
/*Version 0.0.0.5
 * Created on March 26,2013
 * Author: Sneha Dhole
 * Added case special_button
 * Renamed tab case to tabledata
 */
#endregion

namespace Selenium_Framework
{
    public class Selenium_Framework
    {

       private static IWebDriver driver;
        private static FirefoxProfile ffp;
        private string _elogPath = @"c:\webautolog.csv";
        private string _error = "Error in Function or case ";
        private string tagname = null;
        public string actval = null;

        public string eLogPath
        {
            get
            {
                return _elogPath;
            }
            set
            {
                _elogPath = value;
            }
        }

        public enum appBrowser
        {
            IE,
            Firefox,
            Chrome
        };
        private void doubleclick(IWebElement el, IWebDriver dr)
        {

            Actions acn = new Actions(dr);
            acn.DoubleClick(el);
           
            
        }
        /// <summary>
        /// This function is used for logging
        /// </summary>
        /// <param name="spath">Should be path of file</param>
        /// <param name="stxtMsg">Should be message o be logged</param>
        private void logTofile(string spath, string stxtMsg)
        {
            try
            {
                if (!System.IO.File.Exists(spath))
                {
                    System.IO.File.AppendAllText(spath, "Date            Time , Action" + Environment.NewLine);
                }
                else
                {
                    System.IO.File.AppendAllText(spath, System.DateTime.Now + " , " + stxtMsg + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(_error + "[logTofile]:" + Environment.NewLine + ex.Message);
            }

        }

        /// <summary>
        /// To launch the URL on specified browser
        /// </summary>
        /// <param name="browser">Should be value from enum appBrowser</param>
        /// <param name="path">URL of the application</param>
        public void launchweb(appBrowser br, string path)
        {
            logTofile(eLogPath, "In function launchweb");
            switch (br.ToString().ToLower())
            {
                #region Code to launch IE
                case "ie":
                    try
                    {
                        var options = new InternetExplorerOptions();
                        options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                        options.InitialBrowserUrl = path;
                        driver = new InternetExplorerDriver(options);
                        logTofile(eLogPath, "Browser used: IE");
                        driver.Navigate().GoToUrl(path);
                        driver.Manage().Window.Maximize();

                    }
                    catch (Exception ex)
                    {
                        throw new Exception(_error + "[launchweb:IE]:" + Environment.NewLine + ex.Message);
                    }
                    break;
                #endregion

                #region Code to launch firefox
                case "firefox":
                    try
                    {
                        ffp = new FirefoxProfile();
                        ffp.AcceptUntrustedCertificates = true;
                        driver = new FirefoxDriver(ffp);
                        logTofile(eLogPath, "Browser used: Firefox");
                        driver.Navigate().GoToUrl(path);
                        driver.Manage().Window.Maximize();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(_error + "[launchweb:Firefox]:" + Environment.NewLine + ex.Message);
                    }
                    break;
                #endregion

                #region Code to launch chrome
                case "chrome":
                    try
                    {
                        driver = new ChromeDriver();
                        logTofile(eLogPath, "Browser used: Chrome");
                        driver.Navigate().GoToUrl(path);
                        driver.Manage().Window.Maximize();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(_error + "[launchweb:Chrome]:" + Environment.NewLine + ex.Message);
                    }
                    break;
                #endregion
            }
        }

        /// <summary>
        /// This function goes to main page from any frame or iframe
        /// </summary>
        public void goToFirstFrame()
        {
            driver.SwitchTo().DefaultContent();
        }

        /// <summary>
        /// This function is used to perform action on the control type
        /// </summary>
        /// <param name="_controlType">Should be name of control type</param>
        /// <param name="_searchBy">Should be name of control type</param>
        /// <param name="_searchByValue">Should be value/regular expression of Name or Id or Linktext or Index</param>
        /// <param name="_index">Ordinal identifier used when there are more control types of same properties or no properties</param>
        /// <param name="_regexp">Should be Y or N</param>
        /// <param name="_controlValue">Should be value as per the action to be perfomed</param>
        public void perfaction(string _controlType, string _searchBy, string _searchByValue, string _controlValue, int _index = 0, string _regexp = "n")
        {
            try
            {
                logTofile(eLogPath, "In function perfaction");
                switch (_controlType.Trim().ToLower())
                {
                    #region Code for text

                    case "text":
                        try
                        {
                            logTofile(_elogPath, "[Text]: Inside Text case");
                            if (_controlValue.Length > 0)
                            {
                                tagname = "Input";
                                IWebElement webtext = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webtext != null)
                                {
                                    logTofile(eLogPath, "[text]:Found control");
                                    webtext.Clear();
                                    webtext.SendKeys(_controlValue);
                                    logTofile(eLogPath, "[text]:Entered the value :" + _controlValue);
                                }
                                else
                                {
                                    logTofile(eLogPath, "[text]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:text]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for checkbox

                    case "checkbox":
                        try
                        {
                            logTofile(_elogPath, "[Checkbox]: Inside Checkbox case");
                            tagname = "Input";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webcheck = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webcheck != null)
                                {
                                    logTofile(eLogPath, "[checkbox]:Found control");
                                    webcheck.Click();
                                    logTofile(eLogPath, "[checkbox]:changed the state of checkbox");
                                }
                                else
                                {
                                    logTofile(eLogPath, "[checkbox]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:checkbox]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for button

                    case "button":
                        try
                        {
                            logTofile(_elogPath, "[Button]: Inside Button case");
                            tagname = "Input";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webbutton = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webbutton != null)
                                {
                                    logTofile(eLogPath, "[button]:Found control");
                                    webbutton.Click();
                                    logTofile(eLogPath, "[button]:clicked the button");
                                }
                                else
                                {
                                    logTofile(eLogPath, "[button]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:button]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for special button

                    case "special_button":
                        try
                        {
                            logTofile(_elogPath, "[special button]: Inside Button case");
                            tagname = "button";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webspbutton = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webspbutton != null)
                                {
                                    logTofile(eLogPath, "[special button]:Found control");
                                    webspbutton.Click();
                                    logTofile(eLogPath, "[special button]:clicked the button");
                                }
                                else
                                {
                                    logTofile(eLogPath, "[special button]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:special button]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for link

                    case "link":
                        try
                        {
                            logTofile(_elogPath, "[Link]: Inside Link case");
                            tagname = "a";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement weblink = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);


 

                                if (weblink != null)
                                {
                                    logTofile(eLogPath, "[Link]:Found control");
                                    weblink.Click();
                                    logTofile(eLogPath, "[Link]:clicked the Link");
                                }
                                else
                                {
                                    logTofile(eLogPath, "[Link]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:Link]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for dropdown

                    case "dropdown":
                        try
                        {
                            logTofile(_elogPath, "[Dropdown]: Inside dropdown case");
                            tagname = "Select";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webdropdown = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webdropdown != null)
                                {
                                    logTofile(eLogPath, "[Dropdown]:Found control");
                                    SelectElement clickthis = new SelectElement(webdropdown);
                                    clickthis.SelectByText(_controlValue);
                                    logTofile(eLogPath, "[Dropdown]:Selected the element " + _controlValue);
                                }
                                else
                                {
                                    logTofile(eLogPath, "[Dropdown]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:Dropdown]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for image

                    case "img":
                        try
                        {
                            logTofile(_elogPath, "[Image]: Inside img case");
                            tagname = "img";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webimg = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webimg != null)
                                {
                                    logTofile(eLogPath, "[Image]:Found control");
                                    try
                                    {
                                        webimg.Click();
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception("Inside click: " + ex.Message);
                                    }

                                    logTofile(eLogPath, "[Image]:clicked the Image");
                                }
                                else
                                {
                                    logTofile(eLogPath, "[Image]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:Image]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for textarea

                    case "textarea":
                        try
                        {
                            logTofile(_elogPath, "[Textarea]: Inside textarea case");
                            tagname = "textarea";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webtextarea = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webtextarea != null)
                                {
                                    logTofile(eLogPath, "[Textarea]:Found control");
                                    webtextarea.Clear();
                                    webtextarea.SendKeys(_controlValue);
                                    logTofile(eLogPath, "[Textarea]:Entered the text");
                                }
                                else
                                {
                                    logTofile(eLogPath, "[Textarea]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:Textarea]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for Frame

                    case "frame":
                        try
                        {
                            logTofile(_elogPath, "[Frame]: Inside frame case");
                            tagname = "frame";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webframe = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webframe != null)
                                {
                                    logTofile(eLogPath, "[Frame]:Found control");
                                    driver.SwitchTo().Frame(webframe);
                                    logTofile(eLogPath, "[Frame]:Switched to Frame");
                                }
                                else
                                {
                                    logTofile(eLogPath, "[Frame]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:Frame]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for Iframe

                    case "iframe":
                        try
                        {
                            logTofile(_elogPath, "[Iframe]: Inside iframe case");
                            tagname = "iframe";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webiframe = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webiframe != null)
                                {
                                    logTofile(eLogPath, "[Iframe]:Found control");
                                    driver.SwitchTo().Frame(webiframe);
                                    logTofile(eLogPath, "[Iframe]:Switched to Iframe");
                                }
                                else
                                {
                                    logTofile(eLogPath, "[Iframe]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:Iframe]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for Table data

                    case "tabledata":
                        try
                        {
                            logTofile(_elogPath, "[Tabledata]: Inside tabledata case");
                            tagname = "td";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webtd = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webtd != null)
                                {
                                    logTofile(eLogPath, "[Tabledata]:Found control");
                                    webtd.Click();
                                    logTofile(eLogPath, "[Tabledata]:clicked the Tabledata");
                                }
                                else
                                {
                                    logTofile(eLogPath, "[Tabledata]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:Tabledata]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    default:
                        logTofile(_elogPath, "Invalid control type");
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(_error + "[perfaction]:" + Environment.NewLine + ex.Message);
            }

        }

        /// <summary>
        /// This function is used to identify the element
        /// </summary>
        /// <param name="searchBy">Should be Name or Id or Linktext or Index or src or onclick</param>
        /// <param name="searchByValue">Should be value/regular expression of Name or Id or Linktext or Index or src or onclick</param>
        /// <param name="index">Ordinal identifier used when there are more control types of same properties or no properties</param>
        /// <param name="regexp">Should be Y or N</param>
        /// <param name="tagname">Should be Input , Select or a</param>
        /// <returns>Webelement</returns>
        private IWebElement GetWebElement(string searchBy, string searchByValue, int index, string regexp, string tagname)
        {
            logTofile(eLogPath, "In function GetWebElement");
            IWebElement webelement = null;
            string isreg = regexp.ToLower();
            logTofile(eLogPath, "Value of isreg : " + isreg);
            try
            {
                IList<IWebElement> webelements = driver.FindElements(By.TagName(tagname));
                logTofile(_elogPath, "Count of Elements : " + webelements.Count);
                switch (searchBy.ToLower())
                {
                    #region code for name case

                    case "name":


                        try
                        {
                            if (isreg == "n")
                            {
                                logTofile(_elogPath, "searchValue is not a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Name of element " + i + " : " + webelements[i].GetAttribute("Name"));
                                    if (webelements[i].GetAttribute("Name") != null)
                                    {
                                        if (webelements[i].GetAttribute("Name").Contains(searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's name");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found with matching name");
                            }
                            else
                            {
                                logTofile(_elogPath, "searchValue is a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Name of element " + i + " : " + webelements[i].GetAttribute("Name"));
                                    if (webelements[i].GetAttribute("Name") != null)
                                    {
                                        if (Regex.IsMatch(webelements[i].GetAttribute("Name"), searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's name's regular expression");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found having name matching with regular expression");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[GetWebElement:Name]:" + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region code for id case

                    case "id":
                        try
                        {
                            if (isreg == "n")
                            {
                                logTofile(_elogPath, "searchValue is not a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Id of element " + i + " : " + webelements[i].GetAttribute("Id"));
                                    if (webelements[i].GetAttribute("Id") != null)
                                    {
                                        if (webelements[i].GetAttribute("Id").Contains(searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's id");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found with matching name");
                            }
                            else
                            {
                                logTofile(_elogPath, "searchValue is a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Id of element " + i + " : " + webelements[i].GetAttribute("Id"));
                                    if (webelements[i].GetAttribute("Id") != null)
                                    {
                                        if (Regex.IsMatch(webelements[i].GetAttribute("Id"), searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's id's regular expression");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found having id matching with regular expression");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[GetWebElement:Id]:" + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region code for value case

                    case "value":
                        try
                        {
                            if (isreg == "n")
                            {
                                logTofile(_elogPath, "searchValue is not a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Value of element " + i + " : " + webelements[i].GetAttribute("value"));
                                    if (webelements[i].GetAttribute("value") != null)
                                    {
                                        if (webelements[i].GetAttribute("value").Contains(searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's value");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found with matching name");
                            }
                            else
                            {
                                logTofile(_elogPath, "searchValue is a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Id of element " + i + " : " + webelements[i].GetAttribute("Id"));
                                    if (webelements[i].GetAttribute("Id") != null)
                                    {
                                        if (Regex.IsMatch(webelements[i].GetAttribute("Id"), searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's id's regular expression");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found having id matching with regular expression");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[GetWebElement:Id]:" + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region code for Linktext case
                    case "linktext":
                        try
                        {
                            if (isreg == "n")
                            {
                                logTofile(_elogPath, "searchValue is not a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Text of element " + i + " : " + webelements[i].Text);
                                    if (webelements[i].Text != null)
                                    {
                                        if (webelements[i].Text.Contains(searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's text");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found with matching text");
                            }
                            else
                            {
                                logTofile(_elogPath, "searchValue is a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Text of element " + i + " : " + webelements[i].Text);
                                    if (webelements[i].Text != null)
                                    {
                                        if (Regex.IsMatch(webelements[i].Text, searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's text's regular expression");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found having text matching with regular expression");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[GetWebElement:Linktext]:" + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region code for Src case
                    case "src":
                        try
                        {
                            if (isreg == "n")
                            {
                                logTofile(_elogPath, "searchValue is not a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Src of element " + i + " : " + webelements[i].GetAttribute("src"));
                                    if (webelements[i].GetAttribute("src") != null)
                                    {
                                        if (webelements[i].GetAttribute("src").Contains(searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's src");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found with matching src");
                            }
                            else
                            {
                                logTofile(_elogPath, "searchValue is a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "Src of element " + i + " : " + webelements[i].GetAttribute("src"));
                                    if (webelements[i].GetAttribute("src") != null)
                                    {
                                        if (Regex.IsMatch(webelements[i].GetAttribute("src"), searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's src's regular expression");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found having src matching with regular expression");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[GetWebElement:Src]:" + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region code for onclick case
                    case "onclick":
                        try
                        {
                            if (isreg == "n")
                            {
                                logTofile(_elogPath, "searchValue is not a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "onclick of element " + i + " : " + webelements[i].GetAttribute("onclick"));
                                    if (webelements[i].GetAttribute("onclick") != null)
                                    {
                                        if (webelements[i].GetAttribute("onclick").Contains(searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's onclick");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found with matching onclick");
                            }
                            else
                            {
                                logTofile(_elogPath, "searchValue is a regular expression.");
                                int j = 0;
                                for (int i = 0; i < webelements.Count; i++)
                                {
                                    logTofile(_elogPath, "onclick of element " + i + " : " + webelements[i].GetAttribute("onclick"));
                                    if (webelements[i].GetAttribute("onclick") != null)
                                    {
                                        if (Regex.IsMatch(webelements[i].GetAttribute("onclick"), searchByValue))
                                        {
                                            logTofile(_elogPath, "searchValue matched with element's onclick's regular expression");
                                            if (index <= 0)
                                            {
                                                webelement = webelements[i];
                                                logTofile(_elogPath, "webelement found.");
                                                logTofile(_elogPath, "Value of i : " + i);
                                                break;
                                            }
                                            else
                                            {
                                                if (j == index)
                                                {
                                                    webelement = webelements[i];
                                                    logTofile(_elogPath, "webelement found.");
                                                    logTofile(_elogPath, "Value of index : " + j);
                                                    break;
                                                }
                                                else
                                                    logTofile(_elogPath, "Index not matched");
                                            }
                                            j++;
                                        }
                                    }
                                }
                                if (webelement == null)
                                    logTofile(_elogPath, "Element not found having onclick matching with regular expression");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[GetWebElement:onclick]:" + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region code for index case

                    case "index":
                        try
                        {
                            int ind = Convert.ToInt32(searchByValue);
                            for (int i = 0; i < webelements.Count; i++)
                            {
                                if (ind == i)
                                {
                                    logTofile(_elogPath, "index found");
                                    webelement = webelements[i];
                                    logTofile(_elogPath, "webelement found.");
                                    logTofile(_elogPath, "Value of i : " + i);
                                    break;
                                }
                                else
                                    logTofile(_elogPath, "Index not matched");
                            }
                            if (webelement == null)
                                logTofile(_elogPath, "Element not found with matching index");

                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[GetWebElement:Index]:" + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    default:
                        logTofile(_elogPath, "Invalid searchBy value");
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(_error + "[GetWebElement]:" + Environment.NewLine + ex.Message);
            }

            return webelement;
        }

        /// <summary>
        /// This function is to verify the data
        /// </summary>
        /// <param name="_controlType">Should be name of control type</param>
        /// <param name="_searchBy">Should be name of control type</param>
        /// <param name="_searchByValue">Should be value/regular expression of Name or Id or Linktext or Index or src or onclick</param>
        /// <param name="_controlValue">Value to be compared with</param>
        /// <param name="_index">Ordinal identifier used when there are more control types of same properties or no properties</param>
        /// <param name="_regexp">Should be Y or N</param>
        /// <returns></returns>
        /// 


        public string getdata(string _controlType, string _searchBy, string _searchByValue, string _controlValue, int _index = 0, string _regexp = "n")
        {
            try
            {
                logTofile(eLogPath, "In function getdata");
                switch (_controlType.Trim().ToLower())
                {
                    #region Code for text

                    case "text":
                        try
                        {
                            logTofile(_elogPath, "[Text]: Inside Text case");
                            if (_controlValue.Length > 0)
                            {
                                tagname = "Input";
                                IWebElement webtext = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webtext != null)
                                {
                                    logTofile(eLogPath, "[text]:Found control");
                                    actval = webtext.GetAttribute("value");
                                    if (actval != null)
                                    {
                                        logTofile(eLogPath, "[text]:Actual value :" + actval);
                                        if (_controlValue == actval)
                                            logTofile(_elogPath, "text values matches with expected value");
                                        else
                                            logTofile(_elogPath, "text values does not match with expected value");
                                    }
                                    else
                                    {
                                        logTofile(_elogPath, "text value not found");
                                    }
                                }
                                else
                                {
                                    logTofile(eLogPath, "[text]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[getdata:text]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for checkbox

                    case "checkbox":
                        try
                        {
                            logTofile(_elogPath, "[Checkbox]: Inside Checkbox case");
                            tagname = "Input";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webcheck = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webcheck != null)
                                {
                                    logTofile(eLogPath, "[checkbox]:Found control");
                                    actval = webcheck.Selected.ToString();
                                    if (actval != null)
                                    {
                                        logTofile(eLogPath, "[checkbox]:Actual value :" + actval);
                                        if (_controlValue == actval)
                                            logTofile(_elogPath, "checkbox state matches with expected state");
                                        else
                                            logTofile(_elogPath, "checkbox state does not match with expected state");
                                    }
                                    else
                                    {
                                        logTofile(_elogPath, "checkbox state not found");
                                    }
                                }
                                else
                                {
                                    logTofile(eLogPath, "[checkbox]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[perfaction:checkbox]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for dropdown

                    case "dropdown":
                        try
                        {
                            logTofile(_elogPath, "[Dropdown]: Inside dropdown case");
                            tagname = "Select";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webdropdown = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webdropdown != null)
                                {
                                    logTofile(eLogPath, "[Dropdown]:Found control");
                                    SelectElement clickthis = new SelectElement(webdropdown);
                                    actval = clickthis.SelectedOption.Text;
                                    if (actval != null)
                                    {
                                        logTofile(eLogPath, "[dropdown]:Actual value :" + actval);
                                        if (_controlValue == actval)
                                            logTofile(_elogPath, "selected option in dropdown matches with expected option");
                                        else
                                            logTofile(_elogPath, "selected option in dropdown does not match with expected option");
                                    }
                                    else
                                    {
                                        logTofile(_elogPath, "dropdown selected option not found");
                                    }
                                }
                                else
                                {
                                    logTofile(eLogPath, "[Dropdown]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[getdata:Dropdown]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    #region Code for textarea

                    case "textarea":
                        try
                        {
                            logTofile(_elogPath, "[Textarea]: Inside textarea case");
                            tagname = "textarea";
                            if (_controlValue.Length > 0)
                            {
                                IWebElement webtextarea = GetWebElement(_searchBy, _searchByValue, _index, _regexp, tagname);
                                if (webtextarea != null)
                                {
                                    logTofile(eLogPath, "[Textarea]:Found control");
                                    actval = webtextarea.GetAttribute("value");
                                    if (actval != null)
                                    {
                                        logTofile(eLogPath, "[textarea]:Actual value :" + actval);
                                        if (_controlValue == actval)
                                            logTofile(_elogPath, "text values matches with expected value");
                                        else
                                            logTofile(_elogPath, "text values does not match with expected value");
                                    }
                                    else
                                    {
                                        logTofile(_elogPath, "text value not found");
                                    }
                                }
                                else
                                {
                                    logTofile(eLogPath, "[Textarea]:Control not found");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(_error + "[getdata:Textarea]: " + Environment.NewLine + ex.Message);
                        }
                        break;

                    #endregion

                    default:
                        logTofile(_elogPath, "Invalid control type");
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(_error + "[getdata]:" + Environment.NewLine + ex.Message);
            }
            return actval;

        }

        public IWebDriver getDriver(){
            
            return driver;
        }
        
    }
}
