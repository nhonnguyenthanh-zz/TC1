using System;
using OpenQA.Selenium;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using TC1.executionEngine;
using TC1.utility;

namespace TC1.config
{
    class ActionKeywords
    {
        public static IWebDriver driver;
        public static int wait = 2000;
        public static void openBrowser(String object_, String browser)
        {

            try
            {
                if (browser.Equals("Firefox"))
                {
                    driver = new FirefoxDriver();

                }
                else if (browser.Equals("IE"))
                {
                    //Dummy Code, Implement you own code
                    driver = new InternetExplorerDriver();
                    Log.info("IE browser started");
                }
                else if (browser.Equals("Chrome"))
                {
                    //Dummy Code, Implement you own code
                    driver = new ChromeDriver();
                    Log.info("Chrome browser started");
                }
                else
                {
                    driver = new PhantomJSDriver();
                    Log.info("Headless PhantomJS browser started");
                }

                
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(5));
                driver.Manage().Window.Maximize();
            }
            catch (Exception ex)
            {
                Log.info("Not able to open the Browser --- " + ex.Message);
                DriverScript.bResult = false;
            }
        }

        public static void navigate(String object_, String data)
        {
            try
            {
                Log.info("Navigating to URL");
                driver.Navigate().GoToUrl(Constants.URL);
            }
            catch (Exception ex)
            {
                Log.info("Not able to navigate --- " + ex.Message);
                DriverScript.bResult = false;
            }
        }

        public static void click(String object_, String data)
        {
            try
            {
                Log.info("Clicking on " + object_);
                Thread.Sleep(wait);
                IWebElement element = driver.FindElement(RespositoryParser.getObject(object_));
                element.Click();
            }
            catch (Exception ex)
            {
                Log.error("Not able to click " + object_ + " --- " + ex.Message);
                DriverScript.bResult = false;
            }
        }

        public static void input(String object_, String data)
        {
            try
            {
                Log.info("Entering the text in " + object_);
                Thread.Sleep(wait);
                IWebElement element = driver.FindElement(RespositoryParser.getObject(object_));
                element.Clear();
                Thread.Sleep(wait);
                element.SendKeys(data);
            }
            catch (Exception ex)
            {
                Log.error("Not able to Enter " + object_ + "--- " + ex.Message);
                DriverScript.bResult = false;
            }
        }

        public static void select(String object_, String data)
        {

            try
            {
                Log.info("Choose a value in dropdown " + object_);
                Thread.Sleep(wait);
                IWebElement element = driver.FindElement(RespositoryParser.getObject(object_));
                SelectElement select = new SelectElement(element);
                select.SelectByText(data);
            }
            catch (Exception ex)
            {
                // TODO: handle exception
                Log.error("Not able to choose" + object_ + " --- " + ex.Message);
                DriverScript.bResult = false;
            }

        }
        public static void waitFor(String object_, String data)
        {
		try{
                Log.info("Wait for 5 seconds");
                Thread.Sleep(5000);

            }catch(Exception ex){
                Log.error("Not able to Wait --- " + ex.Message);
                DriverScript.bResult = false;
            }
        }

        public static void closeBrowser(String object_, String data)
        {
            try
            {
                Log.info("Closing the browser");
                driver.Quit();
            }
            catch (Exception ex)
            {
                Log.error("Not able to Close the Browser --- " + ex.Message);
                DriverScript.bResult = false;
            }
        }
    }
}
