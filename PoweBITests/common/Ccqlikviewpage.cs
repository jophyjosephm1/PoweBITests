using System;
using System.Threading;
using System.Dynamic;
using System.IO;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Configuration;

namespace PoweBITests.common
{
	public class Ccqlikviewpage
	{
		public IWebDriver Driver { get; }

		public IWebElement CcqlickviewLink
		{
			get {return Driver.FindElement(By.LinkText("Weekly Reporting - PFC TRIAL.qvw"));}
		}

		public IWebElement FinancialyearSearch
		{
			get { return Driver.FindElement(By.XPath("//*[@id='18']/div[2]/div[1]/div")); }
		}

		public IWebElement MonthSearch
		{
			get { return Driver.FindElement(By.XPath("//*[@id='7']/div[2]/div[1]/div")); }
		}

		public IWebElement WeekEndDateSearch
		{
			get { return Driver.FindElement(By.XPath("//*[@id='9']/div[2]/div[1]/div")); }
		}

		public IWebElement ReportDateSearch
		{
			get { return Driver.FindElement(By.XPath("//*[@id='8']/div[2]/div[1]/div")); }
		}

		public IWebElement StateSearch
		{
			get { return Driver.FindElement(By.XPath("//*[@id='14']/div[2]/div[1]/div")); }
		}
		public IWebElement SearchInput
		{
			get { return Driver.FindElement(By.XPath("html/body/div[2]/input")); }
		}


		
		public IWebElement ExportButton
		{
			get { return Driver.FindElement(By.XPath("//*[@id='5']/div[2]/div[1]/div[3]")); }
		}

		

		public IWebElement TabArrow
		{
			get { return Driver.FindElement(By.XPath("//*[@id='Tabrow']/ul/li[2]/a[2]/span")); }
		}

		public IWebElement PowerBICheckingTab
		{
			get { return Driver.FindElement(By.XPath("//a/span[Contains(text(),'Power BI Checking')]")); }
		}

		public Ccqlikviewpage(IWebDriver driver)
		{
			Driver = driver;
		}

		public void GoTo()
		{
			Driver.Navigate().GoToUrl(ConfigurationManager.AppSettings.Get("ccqlikviewUrl"));
		}

		public Ccqlikviewpage CCliCcqlikviewpageLinkClick()
		{
			CcqlickviewLink.Click();
		  return new Ccqlikviewpage(Driver);
		}


		public void PowerBITabClick()
		{
			//for (int i=0;i<4;i++)
			//{
			//	TabArrow.Click();
			//}

			//PowerBICheckingTab.Click();

			IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
			js.ExecuteScript("scroll(3000,0);");
		}

		public void ChooseDates()
		{
			//selecting financial year filter
			FinancialyearSearch.Click();
			SearchInput.SendKeys("2017" + Keys.Enter);
			Thread.Sleep(3000);

			//selecting Month filter
			DateTime now = DateTime.Now;
			var month = now.ToString("MMM");
			MonthSearch.Click();
			Thread.Sleep(4000);
			SearchInput.SendKeys(month+Keys.Enter);
			Thread.Sleep(3000);

			////selecting weekend date filter
			//DateTime baseDate = DateTime.Today;
			//var today = baseDate;
			//var thisWeekStart = baseDate.AddDays(-(int)baseDate.DayOfWeek);
			//var thisWeekEnd = thisWeekStart.AddDays(7).AddSeconds(-1).Day;
			//WeekEndDateSearch.Click();
			//SearchInput.SendKeys(month+" "+thisWeekEnd+Keys.Enter);
			//Thread.Sleep(3000);

			//Selecting report date filter
			var yesterdayDate = now.AddDays(-1).ToString("d");
			ReportDateSearch.Click();
			Thread.Sleep(4000);
			SearchInput.SendKeys(yesterdayDate+Keys.Enter);
			Thread.Sleep(9000);

			//Selecting state filter
			StateSearch.Click();
			Thread.Sleep(3000);
			SearchInput.SendKeys(Keys.Enter);
			Thread.Sleep(6000);
		}

		public void WaitForPagetoLoad()
		{
			WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(50));
			IWebElement myDynamicElement = wait.Until<IWebElement>((d) =>
			{
				try
				{
					return d.FindElement(By.XPath("//*[@id='3']/div[3]/table/tbody/tr/td"));
				}
				catch
				{
					return null;
				}
			});

		}

		//Exporting PowerBIChecking excel data
		public void ExportExcel()
		{
			ExportExcel(ExportButton);
			Thread.Sleep(5000);
			MoveExportedFile("PowerBIChecking");
		}

	   


		//Exporting data in to Excel
		public void ExportExcel(IWebElement exportbutton)
		{
			exportbutton.Click();

		}



		//Moving exported excel file in to a specific location
		public void MoveExportedFile(String type)
		{
			string pattern = "*.xls";
			var dirInfo = new DirectoryInfo(ConfigurationManager.AppSettings.Get("ReportDownloadLocation"));
			var file = (from f in dirInfo.GetFiles(pattern) orderby f.LastWriteTime descending select f).First();
			System.IO.File.Move(ConfigurationManager.AppSettings.Get("ReportDownloadLocation")+ "/" + file, ConfigurationManager.AppSettings.Get("ReportActualLocation")+"/"+ type + ".xls");

		}

		

		
	}
}
