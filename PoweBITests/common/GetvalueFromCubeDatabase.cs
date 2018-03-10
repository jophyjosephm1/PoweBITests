using System;
using System.Data.OleDb;


namespace PoweBITests.common
{
	 public class GetvalueFromCubeDatabase
	{

		    
            
		public string OpenConenction(string sql,string columnIndex)
		{
			try
			{
				OleDbDataReader myReader = null;
				OleDbCommand command;
                string value = null;

				OleDbConnection myConnection =
					new OleDbConnection("Data Source = auperpsvsql12; Initial Catalog = ExecutiveDashboards; Persist Security Info = False; Integrated Security = SSPI; Provider = MSOLAP");
						
				myConnection.Open();
				command = new OleDbCommand(sql,myConnection);

				myReader = command.ExecuteReader();
				while (myReader.Read())
				{
				  value= (myReader[columnIndex] .ToString());

				}

                	
				command.Dispose();
                return value;

            }

			catch (Exception ex)
			{

                return ex.StackTrace;
			}

		}

		


	}
}
