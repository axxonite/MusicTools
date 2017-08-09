using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace SoundLibTool
{
    public static class WebDriverExtensions
    {
        public static IWebElement FindElement(this IWebDriver driver, By by, int timeoutInSeconds)
        {
            try
            {
                if (timeoutInSeconds > 0)
                {
                    var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
                    return wait.Until(drv => drv.FindElement(by));
                }
                return driver.FindElement(by);
            }
            catch (NoSuchElementException e)
            {
                return null;
            }
            catch (WebDriverException e)
            {
                return null;
            }
        }
    }
}