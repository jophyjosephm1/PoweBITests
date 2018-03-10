using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Configuration;


namespace PoweBITests.common
{
	public class ExcelComparison
	{

        Excel.Application xlApp= new Excel.Application();
	    Excel.Workbook xlBook;
		Excel.Worksheet xlSheet;

		

		
		public string GetCqlickValueFromExcel(string type, string columnIndex)
		{
			
			return XlTotals(type, columnIndex);
			
			
		}

        		

		public string XlTotals(string type, string columnIndex)
		{

			xlBook = xlApp.Workbooks.Open(ConfigurationManager.AppSettings.Get("ReportActualLocation")+"/" + type);
			xlApp.Visible = true;
			xlSheet = xlBook.ActiveSheet;
			int nInLastRow = xlSheet.Cells.Find("*", System.Reflection.Missing.Value,
			System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows,
			Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
			
			if (columnIndex == "N")
			{
				string d = xlSheet.UsedRange.Cells[nInLastRow, columnIndex].Value.ToString();
				double total = Math.Round((double.Parse(d)* 100), 1);
				string sTotal = total.ToString();
				CloseXlApplication();
				return sTotal;

			}
			else
			{
				string d = xlSheet.UsedRange.Cells[nInLastRow, columnIndex].Value.ToString();
				double total = Math.Round((double.Parse(d) / 1000), 2);
				string sTotal = total.ToString();
				CloseXlApplication();
				return sTotal;
			}
			
			


		}

		public void CloseXlApplication()
		{

			xlBook.Close();
			xlApp.Quit();
		}
		public void CleanUpDir()
		{

			System.IO.DirectoryInfo di = new DirectoryInfo(ConfigurationManager.AppSettings.Get("ReportActualLocation"+"/"));

			foreach (FileInfo file in di.GetFiles())
			{
				if(file.Name== "PowerBIChecking.xls")
				file.Delete();
				
			}
		}


	}
}


