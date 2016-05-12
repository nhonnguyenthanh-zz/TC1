using System;
using System.IO;
using OpenQA.Selenium;
using System.Text.RegularExpressions;

namespace TC1.config
{
    class RespositoryParser
    {

        public string pathOR;
        public static string[] lines;
        public RespositoryParser(string FilePath)
        {
            this.pathOR = FilePath;
            lines = File.ReadAllLines(pathOR);
            
        }

        public static By getObject(string locatorName)
        {
            string locatorType = null;
            string locatorValue = null;
            By locator = null;
            for (int i = 0; i < lines.Length; i++)
            {
                
                string Object = Regex.Split(lines[i], "::")[0].Trim();
                if (Object == locatorName)
                {
                    locatorType = Regex.Split(lines[i], "::")[1].Trim().ToString();
                    locatorValue = Regex.Split(lines[i], "::")[2].Trim().ToString();
                    switch (locatorType)
                    {
                        case "Id":
                            locator = By.Id(locatorValue);
                            break;
                        case "Name":
                            locator = By.Name(locatorValue);
                            break;
                        case "CssSelector":
                            locator = By.CssSelector(locatorValue);
                            break;
                        case "LinkText":
                            locator = By.LinkText(locatorValue);
                            break;
                        case "PartialLinkText":
                            locator = By.PartialLinkText(locatorValue);
                            break;
                        case "TagName":
                            locator = By.TagName(locatorValue);
                            break;
                        case "Xpath":
                            locator = By.XPath(locatorValue);
                            break;
                    }
                    break;
                } 
            }
            return locator;
         }     
    } 
}
