using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Countdown
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var searchTerm = "milk";
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("disable-infobars");
            var driver = new ChromeDriver(options); // new instance of browser chrome
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10)); //explicit wait time
            driver.Navigate().GoToUrl("https://shop.countdown.co.nz/"); //navigate to website
            driver.Manage().Window.Maximize(); //maximise chrome browser window
            wait.Until(ExpectedConditions.ElementToBeClickable(driver.FindElementById("search"))).SendKeys(searchTerm);
            driver.FindElementById("searchIcon").Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("search-content")));
            Screenshot searchResultScreenShot = ((ITakesScreenshot)driver).GetScreenshot(); // take screenshot
            WriteToWord.CreateDocument(@"c:\Users\Renju\Documents\Screenshot" + DateTime.Now.Ticks.ToString() + ".docx", searchTerm, searchResultScreenShot.AsByteArray);
            driver.Quit();
        }
    }
}
