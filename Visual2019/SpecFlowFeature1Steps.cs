using System;
using TechTalk.SpecFlow;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using NUnit.Framework;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using NUnit.Framework.Internal;


namespace Visual2019
{
    [Binding]
    public class DecafyingLibrariySteps
    {
        ChromeOptions options = new ChromeOptions();

        IWebDriver driver;// new ChromeDriver();


        [Given(@"I have launched team' site")]
        public void GivenIHaveLaunchedTeamSite()
        {
            _Login();
        }

        [When(@"I select documents library in site contents")]
        public void WhenISelectDocumentsLibraryInSiteContents()
        {
            //Clicks settings button
            IWebElement settingsButton = ObjWebElement("O365_MainLink_Settings", 3);
            settingsButton.Click();
            Thread.Sleep(5000);

            //Go to site contents
            IWebElement siteContentsButton = ObjWebElement("SuiteMenu_zz6_MenuItem_ViewAllSiteContents", 3);
            siteContentsButton.Click();
            Thread.Sleep(10000);
        }

        [Then(@"Document Library is launched")]
        public void ThenDocumentLibraryIsLaunched()
        {

            //Launch document library
            IWebElement documentsLibrary = ObjWebElement("od-FieldRenderer-Renderer-withMetadata", 2, "Documents");
            documentsLibrary.Click();
            Thread.Sleep(20000);


            //Verify two tabs are open
            Assert.AreEqual(2, driver.WindowHandles.Count);          
           
            //Close First tab
            driver.Close();

            //Keep the window handlers in arrays and then manipulate the array
            var mainWindowHandle = driver.WindowHandles[0];
            Assert.IsTrue(!string.IsNullOrEmpty(mainWindowHandle));

            //Close second tab
            driver.SwitchTo().Window(mainWindowHandle);
            driver.Close();

            Thread.Sleep(10000);

        }


        #region
        public void _Login()
        {
            //ChromeDriverSetup
            options.AddArgument("--disable-extensions");
            options.AddArgument("--start-maximized");
           
            //options.AddAdditionalCapability("useAutomationExtension", false);

            //Launch URL in Chrome
            driver = new ChromeDriver(options);
            ClearBrowserCache();
            driver.Manage().Window.Maximize();
            Thread.Sleep(5000);
            driver.Navigate().GoToUrl("https://dongenergyit.sharepoint.com/sites/PRAJATEAMS");
            Thread.Sleep(40000);


            //Enter Username
            IWebElement enterUserName = driver.FindElement(By.Id("i0116"));
            enterUserName.SendKeys("praja@de-mosstest.dk");
            Thread.Sleep(5000);

            //Next Button
            IWebElement nextButton = driver.FindElement(By.Id("idSIButton9"));
            Assert.IsTrue(nextButton.Displayed);
            nextButton.Click();
            Thread.Sleep(10000);

            //Enter Password
            IWebElement password = ObjWebElement("i0118", 3);
            //To be encryted 
            password.SendKeys("Dong1@rsted");

            //Sign In
            IWebElement signIn = ObjWebElement("idSIButton9", 3);
            signIn.Click();
            Thread.Sleep(10000);

            //Stay Signed In
            IWebElement staySignedIn = ObjWebElement("idBtn_Back", 3);
            staySignedIn.Click();
            Thread.Sleep(40000);

        }

        public void ClearBrowserCache()
        {
            driver.Manage().Cookies.DeleteAllCookies(); //delete all cookies
            Thread.Sleep(5000); //wait 5 seconds to clear cookies.

        }

        private void TakeScreenShots()
        {
            string fileName = $"{DateTime.Now.ToString("HHmmss")}.jpeg";
            //string fileName = $"{SnapShotPath + DateTime.Now.ToString("HHmmss")}.jpeg";
            //string fileName = @"C:\Users\WAAZI\source\repos\VITAL\TestResults\" + DateTime.Now.ToString("HHmmss") + ".jpeg";
            //string fileName = @"H:\New folder\TestResults" + DateTime.Now.ToString("HHmmss") + ".jpeg";
            Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
            sc.SaveAsFile(fileName, ScreenshotImageFormat.Jpeg);
        }

        private IWebElement ObjWebElement(string strWebElement, int iType = 0, string textTofind = "", bool exactWord = true)
        {
            if (iType == 0)
                return driver.FindElement(By.XPath(strWebElement));

            else if (iType == 1)
            {
                return driver.FindElement(By.ClassName(strWebElement));
            }

            else if (iType == 2)
            {
                var allElement = driver.FindElements(By.ClassName(strWebElement));
                if (allElement.Count == 0)
                {
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                    allElement = driver.FindElements(By.ClassName(strWebElement));
                }
                foreach (IWebElement t in allElement)
                {
                    if (exactWord)
                    {
                        if (t.Text.ToUpper() == textTofind.ToUpper())
                            return t;
                    }
                    else
                    {
                        if (t.Text.Contains(textTofind))
                            return t;
                    }

                }
            }

            else if (iType == 3)
            {
                return driver.FindElement(By.Id(strWebElement));
            }

            else if (iType == 4)
            {
                return driver.FindElement(By.Name(strWebElement));
            }

            else if (iType == 5)
            {
                return driver.FindElement(By.LinkText(strWebElement));
            }
            //return driver.FindElement(By.ClassName(strWebElement));
            return null;
        }


        private IWebElement tagNameElement(string tagName, string elementText, bool exactWord = true)
        {

            var tagElement = driver.FindElements(By.TagName(tagName));

            if (tagElement.Count == 0)
            {
                Thread.Sleep(5000);
                tagElement = driver.FindElements(By.TagName(tagName));

            }
            foreach (IWebElement t in tagElement)
            {
                if (exactWord)
                {
                    if (t.Text.ToUpper() == elementText.ToUpper())
                        return t;
                }
                else
                {
                    if (t.Text.Contains(elementText))
                        return t;
                }
            }

            return null;
            //return driver.FindElement(By.Id(strFrameId));

        }

        private IWebElement iFrameElement(string classElement, string elementId)
        {

            var frameElement = driver.FindElements(By.ClassName(classElement));

            if (frameElement.Count == 0)
            {
                Thread.Sleep(5000);
                frameElement = driver.FindElements(By.ClassName(classElement));

            }
            //var strFrameId = "";
            string strFrameId;
            foreach (IWebElement t in frameElement)
            {
                strFrameId = t.GetAttribute("Id");
                driver.SwitchTo().Frame(strFrameId);

                IWebElement frameElement1 = driver.FindElement(By.Id(elementId));
                if (frameElement1 != null)
                    return frameElement1;
            }

            return null;
            //return driver.FindElement(By.Id(strFrameId));

        }

        private IWebElement cssSelectorElement(string strWebElement, string textTofind = "", bool exactWord = true)
        {
            var allElement = driver.FindElements(By.CssSelector(strWebElement));

            foreach (IWebElement t in allElement)
            {
                if (exactWord)
                {
                    if (t.Text.ToUpper() == textTofind.ToUpper())
                        return t;
                }
                else
                {
                    if (t.Text.Contains(textTofind))
                        return t;
                }

            }

            return null;

        }

        #endregion

    }
}
