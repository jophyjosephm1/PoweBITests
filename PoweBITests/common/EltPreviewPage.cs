
using OpenQA.Selenium;
using System.Threading;
using System.Configuration;
namespace PoweBITests.common

{
	
	public class EltPreviewPage:GetvalueFromCubeDatabase
	{
        string sqlQuery = null;


        public string getTheEltPageValue(string valueType)
        {
           string value= OpenConenction(getTheSqlQuery(valueType), getTheColumnIndex(valueType));
            return value;
            
        }

        public string getTheSqlQuery(string valueType)
        {
            switch (valueType)
            {
                case "CaAdavance_Daily":
                {
                     return   sqlQuery = @" DEFINE   VAR __DS0FilterTable =      FILTER(KEEPFILTERS(VALUES('CCPOS Store'[StoreType])), 'CCPOS Store'[StoreType] = ""Corporate"")
			        VAR __DS0FilterTable2 = FILTER(KEEPFILTERS(VALUES('Date'[FutureDateFlag])), 'Date'[FutureDateFlag] = 0)
			        VAR __DS0FilterTable3 = FILTER(KEEPFILTERS(VALUES('Date'[Date])), 'Date'[Date] = DATE(YEAR(NOW()-2),Month(NOW()-2),Day(NOW()-2)))
			        VAR __DS0FilterTable4 = FILTER(KEEPFILTERS(VALUES('PeriodMeasureSelect'[Period])), 'PeriodMeasureSelect'[Period] = ""Daily"")
			        EVALUATE TOPN(     1002,     SUMMARIZECOLUMNS('Date'[FiscalYearMonth], __DS0FilterTable, __DS0FilterTable2, __DS0FilterTable3, __DS0FilterTable4, ""CA_Advances_PTD"", 'CCPF Advances'[CA Advances PTD], ""CA_Advances_Budget_PTD"", 'CCPOS Budgets'[CA Advances Budget PTD]),     'Date'[FiscalYearMonth],     0   )  
			        ORDER BY   'Date'[FiscalYearMonth]";
              }

                default:
                    return "Invalid query type";

            }


        }


        public string getTheColumnIndex(string valueType)
        {

            string columnIndex = null;
            switch (valueType)
            {
                case "CaAdavance_Daily":
                    {
                        return columnIndex = "[CA Advances PTD]";
                    }

                default:
                    return "Invalid value type";

            }


        }





    }
}
