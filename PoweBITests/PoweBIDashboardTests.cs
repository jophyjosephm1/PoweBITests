using System;
using System.IO;
using System.Threading;
using NUnit.Framework;
using RelevantCodes.ExtentReports;
using System.Configuration;
using System.Drawing.Text;
using NUnit.Framework.Interfaces;
using PoweBITests.common;
using OpenQA.Selenium;

namespace PoweBITests 
{
	[TestFixture]
	[Description("Compare CCqlickview  and Elt preview pages")]
	public class PoweBIDashboardTests:BaseClass
	{
		
		private Ccqlikviewpage ccqlickviewpage => new Ccqlikviewpage(Driver);
		private EltPreviewPage eltPreviewPage => new EltPreviewPage();
		private ExcelComparison excelComparison=> new ExcelComparison();
		private GetvalueFromCubeDatabase getDbvalues => new GetvalueFromCubeDatabase();
        ExtentReports extent;
		ExtentTest test;
       
        

        //  Initialising the extent report

        [OneTimeSetUp]
	
       public void StartReport()
        {
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actualPath = path.Substring(0, path.LastIndexOf("bin"));
            string projectPath = new Uri(actualPath).LocalPath;
            string reportPath = projectPath + "Reports\\Test.html";
            extent = new ExtentReports(reportPath, true);
        }


		

		/*
	   ****************************************************
		Extracting  CCqlick report and dowloading it locally.
       *****************************************************
        */


		[Test,Order(1)]
		public void ExtractExcel_from_CcqlikView()
		{
			test = extent.StartTest("ExtractExcel_from_CcqlikView");
            InitializeDriver();
            ccqlickviewpage.GoTo();
			Thread.Sleep(5000);
			ccqlickviewpage.CCliCcqlikviewpageLinkClick();
			ccqlickviewpage.WaitForPagetoLoad();
			Thread.Sleep(5000);
			ccqlickviewpage.ChooseDates();
			Thread.Sleep(7000);
			ccqlickviewpage.PowerBITabClick();
			ccqlickviewpage.ExportExcel();
			//Checking file exists
			var fileName = ConfigurationManager.AppSettings.Get("ReportActualLocation")+"/"+"PowerBIChecking.xls";
			Assert.IsTrue(File.Exists(fileName));
			test.Log(LogStatus.Pass, "Test Passed");
			
			
		}

        /*
		************************************************	
		Comparing CCqlick reprot values  with PowerBI.
		 *************************************************
		*/

        //[Test, Order(2)]

        //public void VerifyTotalRetailSales()
        //{
        //	test = extent.StartTest("VerifyTotalRetailSales");

        //	eltPreviewPage.LaunchEltPreviewpage();
        //       Thread.Sleep(8000);
        //	// CCqlickview  values
        //	var cCqlick_TotalRetailSales = excelComparison.GetCqlickValueFromExcel("PowerBIChecking.xls","O");

        //	//ELT preview values
        //	var Elt_TotalRetailSales = eltPreviewPage.GetRetailsTotal();

        //	Assert.AreEqual(cCqlick_TotalRetailSales, Elt_TotalRetailSales);
        //	test.Log(LogStatus.Pass,"Test passed" );

        //}

        //[Test, Order(3)]
        //public void VerifyWebshopSales()
        //{
        //	test = extent.StartTest("VerifyWebshopSales");
        //	eltPreviewPage.LaunchEltPreviewpage();
        //	Thread.Sleep(8000);

        //	// CCqlickview  values
        //	var cCqlick_WebShopSales = excelComparison.GetCqlickValueFromExcel("PowerBIChecking.xls","C");

        //	//ELT preview values
        //	var Elt_WebshopSales = eltPreviewPage.GetWebShopTotal();

        //	Assert.AreEqual(cCqlick_WebShopSales, Elt_WebshopSales);
        //	test.Log(LogStatus.Pass, "Test passed");

        //}

        //[Test, Order(4)]

        //public void VerifyTotalGpPercentage()
        //{
        //	test = extent.StartTest("VerifyTotalGpPercentage");
        //	eltPreviewPage.LaunchEltPreviewpage();

        //	Thread.Sleep(8000);

        //	// CCqlickview  values
        //	var cCqlick_TotalGPercentage = excelComparison.GetCqlickValueFromExcel("PowerBIChecking.xls","N");

        //	//ELT preview values
        //	var Elt_GpPercentage = eltPreviewPage.GetGpPercentage();

        //	Assert.AreEqual(cCqlick_TotalGPercentage,Elt_GpPercentage);
        //	test.Log(LogStatus.Pass, "Test passed");


        //}

        //[Test, Order(5)]

        //public void VerifyPawnBrokingPercentage()
        //{
        //	test = extent.StartTest("VerifyPawnBrokingPercentage");

        //				// CCqlickview  values
        //	var cCqlick_PawnBrokingInterest = excelComparison.GetCqlickValueFromExcel("PowerBIChecking","I");


        //	//ELT preview values
        //	var Elt_PawnBrokingEnterest = eltPreviewPage


        //	Assert.AreEqual(cCqlick_PawnBrokingInterest, Elt_PawnBrokingEnterest);
        //	test.Log(LogStatus.Pass, "Test passed");

        //}


        ///*
        //  ******************************************	
        //Comparing CCPF reports with PowerBI values.
        //     ********************************************
        //      */

        //[Test, Order(6)]

        //public void VerifyCCPF_BadDebtWrittenOffAmount()
        //{
        //	test = extent.StartTest("VerifyCCPF_BadDebtWrittenOffAmount");
        //	//eltPreviewPage.LaunchEltPreviewpage();

        //	//.Sleep(10000);

        //	// CCqlickview  values
        //	var cCqlick_BadWriteoffAmount = excelComparison.GetCqlickValueFromExcel("Daily-Gross Bad Debt.xlsx", "K");


        //	//ELT preview values
        //	var Elt_BadWriteoffAmount = eltPreviewPage.GetBadDebtWrittenoff();


        //	Assert.AreEqual(cCqlick_BadWriteoffAmount, Elt_BadWriteoffAmount);
        //	test.Log(LogStatus.Pass, "Test passed");

        //}



        [Test, Order(7)]

        public void VerifyCCPF_CaAdavanceAmount()
        {
            test = extent.StartTest("VerifyCCPF_CaAdavanceAmount");
           

            // CCqlickview  values
           //ar cCqlick_CaAdvanceAmount = excelComparison.GetCqlickValueFromExcel("Daily-CAAdvances.xlsx", "J");
            
            //ELT preview values
            var Elt_CaAdavnceAmount = eltPreviewPage.getTheEltPageValue("CaAdavance_Daily");


            Assert.AreEqual(33,Elt_CaAdavnceAmount);
            test.Log(LogStatus.Pass, "Test passed");

        }

        //[Test, Order(8)]

        //public void VerifyCCPF_PlAdavanceAmount()
        //{
        //	test = extent.StartTest("VerifyCCPF_PlAdavanceAmount");
        //	eltPreviewPage.LaunchEltPreviewpage();

        //	Thread.Sleep(10000);

        //	// CCqlickview  values
        //	var cCqlick_PlAdvanceAmount = excelComparison.GetCqlickValueFromExcel("Daily-PLAdvances.xlsx", "J");


        //	//ELT preview values
        //	var Elt_PlAdavnceAmount = eltPreviewPage.GetPlAdavances();


        //	Assert.AreEqual(cCqlick_PlAdvanceAmount, Elt_PlAdavnceAmount);
        //	test.Log(LogStatus.Pass, "Test passed");

        //}


        //[Test, Order(9)]

        //public void VerifyCCPF_MaccAdavanceAmount()
        //{
        //	test = extent.StartTest("VerifyCCPF_MaccAdavanceAmount");
        //	eltPreviewPage.LaunchEltPreviewpage();

        //	Thread.Sleep(10000);

        //	// CCqlickview  values
        //	var cCqlick_MaccAdvanceAmount = excelComparison.GetCqlickValueFromExcel("Daily-MACCAdvances.xlsx", "J");


        //	//ELT preview values
        //	var Elt_MaccAdavnceAmount = eltPreviewPage.GetMaccAdavances();


        //	Assert.AreEqual(cCqlick_MaccAdvanceAmount, Elt_MaccAdavnceAmount);
        //	test.Log(LogStatus.Pass, "Test passed");

        //}





        [TearDown]

		//Getting the  Stack trace of failure tests.
		public void GetResult()
		{
			var status = TestContext.CurrentContext.Result.Outcome.Status;
			var stackTrace = "<pre>" + TestContext.CurrentContext.Result.StackTrace + "</pre>";
			var errorMessage = TestContext.CurrentContext.Result.Message;

			if (status == TestStatus.Failed)
			{
				//string screenShotPath = GetScreenShot.Capture(driver, "ScreenShotName");
				test.Log(LogStatus.Fail, stackTrace + errorMessage);
				//test.Log(LogStatus.Fail, "Snapshot below: " + test.AddScreenCapture(screenShotPath));
			}

			extent.EndTest(test);

			if (Driver == null)
			{
				return;
			}

			Driver.Quit();

		}




		[OneTimeTearDown]
        public void CloseTest()
		{
            extent.Flush();
            extent.Close();
			//DateTime now=DateTime.Now;
			//System.IO.File.Copy(@"C:\Git\Source\Automation\PoweBITests\PoweBITests\Reports\Test.html", @"C:\Git\Source\Automation\PoweBITests\PoweBITests\C:\Git\Source\Automation\PoweBITests\PoweBITests\TestResults-Old"+"/"+"TestResult"+"_"+now+".html");
			//removing downloaded excel files.
			//excelComparison.CleanUpDir();
		}




	}
}
