using System;
using System.Diagnostics;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace PADjpAppSharp
{
    public class BrowserFixture : IDisposable
    {
        public BrowserFixture()
        {
            var opts = new ChromeOptions();
            opts.AddArguments("--incognito");

            // Comment this out if you want to watch or interact with the browser (e.g. for debugging)
            if (!Debugger.IsAttached)
            {
                opts.AddArgument("--headless");
            }

            Driver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), opts, TimeSpan.FromSeconds(10));
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        }

        public IWebDriver Driver { get; set; }

        public void Dispose()
        {
            Driver?.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}
