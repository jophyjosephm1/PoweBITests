
using System;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

using System.Configuration;


namespace PoweBITests.common
{
	public class BaseClass
	{
		
		protected  IWebDriver Driver { get; private set; }
		
		public void InitializeDriver()
		{
     //       Driver = new ChromeDriver(ConfigurationManager.AppSettings.Get("ChromeDriverLocation"));
    //        Driver.Manage().Window.Maximize();
         ChromeOptions option = new ChromeOptions();
         option.AddArgument("--headless");
         Driver = new ChromeDriver(ConfigurationManager.AppSettings.Get("ChromeDriverLocation"),option);
		}

	}


}
