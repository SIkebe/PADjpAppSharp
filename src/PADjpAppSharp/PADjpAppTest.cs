using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using Xunit;

namespace PADjpAppSharp
{
    public class PADjpAppTest : IClassFixture<BrowserFixture>, IClassFixture<WinAppDriverFixture>
    {
        private readonly BrowserFixture _browserFixture;
        private readonly WindowsDriver<WindowsElement> _session;

        public PADjpAppTest(BrowserFixture browserFixture, WinAppDriverFixture winAppDriverFixture)
        {
            _browserFixture = browserFixture;
            _session = winAppDriverFixture.Session;
        }

        [Fact]
        public void HandsOnMain()
        {
            var orderData = ReadOrderData();
            var msnData = ReadMsnData();

            Login();

            var menuHandle = _session.WindowHandles[0];
            RegisterOrders(orderData, msnData, menuHandle);

            // Close the menu form
            _session.SwitchTo().Window(menuHandle);
            _session.FindElementByName("閉じる").Click();
        }

        private void RegisterOrders(List<DataFromExcel> orderData, List<DataFromMsn> msnData, string menuHandle)
        {
            foreach (var order in orderData)
            {
                // Switch to the menu form
                _session.SwitchTo().Window(menuHandle);
                _session.FindElementByName("申請").Click();

                // Switch to the new order form
                _session.SwitchTo().Window(_session.WindowHandles.Single(h => h != menuHandle));

                _session.FindElementByAccessibilityId("PART_TextBox").SendKeys(order.OrderDate.Replace(" 0:00:00", string.Empty));
                _session.FindElementByAccessibilityId("KaishaCD").SendKeys(order.CompanyCode);
                _session.FindElementByAccessibilityId("Kaishamei").SendKeys(order.CompanyName);
                _session.FindElementByAccessibilityId("Shohinmei").SendKeys(order.ProductName);
                _session.FindElementByAccessibilityId("Kakaku").SendKeys(order.Price.ToString());

                _session.FindElementByAccessibilityId("Tani").Click();
                _session.FindElementByName(order.Unit).Click();

                // 申請データの単位をキーにしてMSNのデータを検索し、一致したものの円換算金額を入力
                _session.FindElementByAccessibilityId("Enkakau").SendKeys(msnData.Single(m => m.Unit == order.Unit).Yen.ToString());

                _session.FindElementByName("登録").Click();
                new Actions(_session).SendKeys(Keys.Enter).Perform();
                new Actions(_session).SendKeys(Keys.Enter).Perform();
                _session.FindElementByName("閉じる").Click();
            }
        }

        private void Login()
        {
            _session.FindElementByAccessibilityId("UserId").SendKeys("12345");
            _session.FindElementByAccessibilityId("Password").SendKeys("54321");
            _session.FindElementsByClassName("TextBlock").Single(e => e.Text == "ログイン").Click();
        }

        private List<DataFromMsn> ReadMsnData()
        {
            _browserFixture.Driver.Navigate().GoToUrl("https://www.msn.com/ja-jp/money/currencyconverter");
            var mcrows = _browserFixture.Driver.FindElements(By.ClassName("mcrow"));

            var dataFromMsn = new List<DataFromMsn>();
            foreach (var row in mcrows)
            {
                dataFromMsn.Add(new DataFromMsn(
                    row.FindElement(By.ClassName("cntrycol")).FindElement(By.ClassName("truncated-string")).Text,
                    row.FindElement(By.ClassName("pricecol")).FindElement(By.ClassName("truncated-string")).Text)
                );
            }

            return dataFromMsn;
        }

        private static List<DataFromExcel> ReadOrderData()
        {
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "申請データ.xlsx");
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.Single(s => s.Name == "申請");

            var orderData = new List<DataFromExcel>();
            var row = 2;
            while (!worksheet.Cell($"A{row}").IsEmpty())
            {
                orderData.Add(
                    new DataFromExcel(
                        worksheet.Cell($"A{row}").GetString(),
                        worksheet.Cell($"B{row}").GetString(),
                        worksheet.Cell($"C{row}").GetString(),
                        worksheet.Cell($"D{row}").GetString(),
                        worksheet.Cell($"E{row}").GetString(),
                        worksheet.Cell($"F{row}").GetString()
                    ));

                row++;
            }

            return orderData;
        }
    }

    record DataFromExcel(string OrderDate, string CompanyCode, string CompanyName, string ProductName, string Price, string Unit);
    record DataFromMsn(string Unit, string Yen);
}
