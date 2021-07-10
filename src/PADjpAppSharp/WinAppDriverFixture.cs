using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.IO;

namespace PADjpAppSharp
{
    public class WinAppDriverFixture : IDisposable
    {
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";

        public WinAppDriverFixture()
        {
            var appiumOptions = new AppiumOptions();
            appiumOptions.AddAdditionalCapability("app", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "PADjpApp.exe"));
            appiumOptions.AddAdditionalCapability("deviceName", "WindowsPC");
            Session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appiumOptions);
        }

        public WindowsDriver<WindowsElement> Session { get; set; }

        public void Dispose()
        {
            Session?.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}
