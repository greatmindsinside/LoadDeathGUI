using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Data;
using System.Windows.Forms;
using static LoadDeathsGUI.DeathDivorceForm1;

namespace LoadDeathsGUI
{
    public static class LoadSQL 
    {
        public static void OpenSqlConnection(string sTheSSN)
        {
            //Create the Data Table so we can load the Death data into it
            DataTable dt = new DataTable();
            DataTable CampaignTable = new DataTable();

            //Ptint The SQL Query Being Used...
            //Console.WriteLine(GetSQLQueryForSSN(Submit.sTheSSN));

            using (SqlConnection connection = new SqlConnection(GetConnectionString()))
            {
                //SqlCommand command = new SqlCommand(GetSQLQueryForSSN(Global.sTheSSN), connection);
                SqlCommand command = new SqlCommand(GetSQLQueryForSSN(Global.sTheSSN), connection);
                
                StringBuilder errorMessages = new StringBuilder();

                //The amount of time it tries to run a SQL query before returning an error 
                command.CommandTimeout = 90;
               
                try
                {
                    connection.Open();
                    Console.WriteLine("State: {0}", connection.State);
                    Console.WriteLine("ConnectionString: {0}", connection.ConnectionString);
                }
                catch (SqlException ex)
                {
                    for (int i = 0; i < ex.Errors.Count; i++)
                    {
                        errorMessages.Append("Index #" + i + "\n" +
                            "Message: " + ex.Errors[i].Message + "\n" +
                            "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                            "Source: " + ex.Errors[i].Source + "\n" +
                            "Procedure: " + ex.Errors[i].Procedure + "\n");
                    }
                    MessageBox.Show(errorMessages.ToString(), "Death\\Divorce Loader", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                    Console.WriteLine(errorMessages.ToString());

                    //Need to return control to form1 
                    command.Dispose();
                    return;
                }

                SqlDataAdapter ds = new SqlDataAdapter(command);
                try
                {
                    //Place SQL Data Into Table
                    //TableNum: the number of rows successfully added to or refreshed in the Datatable.
                    int TableNum = ds.Fill(dt);
                    if (TableNum < 1)
                    {
                        MessageBox.Show("SQL: No Rows Found.");
                        return;
                    } else if (TableNum > 1) 
                    {

                        //There are multiple GUIDS
                        //Create a new form window with a drop down with the GUIDS 
                        //They select the GUID then we change the row containing that GUID
                        //MessageBox.Show("There are multiple GUIDS");

                        //FindGUIDs(dt);
                    
                        //SqlCommand command2 = new SqlCommand(GetSQLQueryForCampaignName(Global.aCampgainSegments), connection);
                        //SqlDataAdapter ds2 = new SqlDataAdapter(command2);
                        //try
                        //{
                        //    ds2.Fill(CampaignTable);
                        //}
                        //catch (SqlException ex)
                        //{
                        //    for (int i = 0; i < ex.Errors.Count; i++)
                        //    {
                        //        errorMessages.Append("Index #" + i + "\n" +
                        //            "Message: " + ex.Errors[i].Message + "\n" +
                        //            "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                        //            "Source: " + ex.Errors[i].Source + "\n" +
                        //            "Procedure: " + ex.Errors[i].Procedure + "\n");
                        //    }
                        //    MessageBox.Show(errorMessages.ToString(), "Death\\Divorce Loader", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                        //    Console.WriteLine(errorMessages.ToString());
                        //    //Need to return control to form1 when it can't connect it fucking fails and crashes

                        //    Console.WriteLine("Closing Adaptor due to error...");
                        //    ds2.Dispose();
                        //    return;
                        //}
                        
                        //FindGUIDNames(CampaignTable);

                        //ds2.Dispose();

                        //Form2 frm = new Form2();
                        //frm.ShowDialog();

                        //return;
                    }

                    //Populates an array with the headernames
                    GetHeaderNames(dt);

                }
                catch (SqlException ex)
                {
                    for (int i = 0; i < ex.Errors.Count; i++)
                    {
                        errorMessages.Append("Index #" + i + "\n" +
                            "Message: " + ex.Errors[i].Message + "\n" +
                            "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                            "Source: " + ex.Errors[i].Source + "\n" +
                            "Procedure: " + ex.Errors[i].Procedure + "\n");
                    }
                    MessageBox.Show(errorMessages.ToString(), "Death\\Divorce Loader", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                    Console.WriteLine(errorMessages.ToString());
                    //Need to return control to form1 when it can't connect it fucking fails and crashes

                    Console.WriteLine("Closing Adaptor due to error...");
                    ds.Dispose();
                    return;
                }



                //Console.WriteLine("Closing Adaptor...");
                ds.Dispose();
                

                //Console.WriteLine("Passing Table to OpenExcel function...");
                CreateExcelWorksheet.OpenExcel(dt);

            }
        }


        static public void CheckForSQLOpenErrors()
        {

        }

       static public void SQLAdaptor(SqlDataAdapter ds, DataTable dt)
        {

            StringBuilder errorMessages = new StringBuilder();

            try
            {
                //Place SQL Data Into Table
                //TableNum: the number of rows successfully added to or refreshed in the Datatable.
                int TableNum = ds.Fill(dt);
                if (TableNum < 1)
                {
                    MessageBox.Show("SQL: No Rows Found.");
                    return;
                }
                else if (TableNum > 1)
                {

                    //There are multiple GUIDS
                    //Create a new form window with a drop down with the GUIDS 
                    //They select the GUID then we change the row containing that GUID
                    MessageBox.Show("There are multiple GUIDS");

                    FindGUIDs(dt);
                    Form2 frm = new Form2();
                    frm.ShowDialog();
                }

                //Populates an array with the headernames
                GetHeaderNames(dt);

            }
            catch (SqlException ex)
            {
                for (int i = 0; i < ex.Errors.Count; i++)
                {
                    errorMessages.Append("Index #" + i + "\n" +
                        "Message: " + ex.Errors[i].Message + "\n" +
                        "LineNumber: " + ex.Errors[i].LineNumber + "\n" +
                        "Source: " + ex.Errors[i].Source + "\n" +
                        "Procedure: " + ex.Errors[i].Procedure + "\n");
                }
                MessageBox.Show(errorMessages.ToString(), "Death\\Divorce Loader", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                Console.WriteLine(errorMessages.ToString());
                //Need to return control to form1 when it can't connect it fucking fails and crashes

                Console.WriteLine("Closing Adaptor due to error...");
                ds.Dispose();
                return;
            }
        }

       static private string GetConnectionString()
        {
            // To avoid storing the connection string in your code.   
            string ConnString = "Data Source=prodSqlAg.extendhealth.com;" +
            "Initial Catalog=ProdExtendHealth;" +
            "Integrated Security=SSPI;" +
            "Persist Security Info=True;";

            return ConnString;
        }

       static private string SafeGetString(this SqlDataReader reader, int colIndex)
        {
            //Not using this function
            if (!reader.IsDBNull(colIndex))
            {
                return reader.GetString(colIndex);
            }
            return "NULL";
        }

       private static void ReadSingleRow(SqlDataReader reader, IDataRecord record)
        {
            int iTotalCol = record.FieldCount;
            Console.WriteLine(iTotalCol);
            
            for (int i = 0; i < iTotalCol; i++)
            {

                //Get headers
                Console.WriteLine(reader.GetName(i));
                    

                //Console.WriteLine(reader.IsDBNull(i));
                if (!reader.IsDBNull(i))
                {
                    Console.WriteLine(record[i]);
                }
                else
                {
                    Console.WriteLine("");
                }




            }

          
        }

       private static List<string> FindGUIDs(DataTable TheDataTable)
        {

            List<string> campaginSegment = new List<string>();
            DataRow[] rows = TheDataTable.Select();
            for (int i = 0; i < rows.Length; i++)
            {
                campaginSegment.Add(rows[i]["campaignSegmentGuid"].ToString());
                Console.WriteLine(rows[i]["campaignSegmentGuid"]);
            }

            return Global.aCampgainSegments = campaginSegment;
        }

        private static List<string> FindGUIDNames(DataTable TheDataTable)
        {

            List<string> campaginSegment = new List<string>();
            DataRow[] rows = TheDataTable.Select();
            for (int i = 0; i < rows.Length; i++)
            {
                campaginSegment.Add(rows[i]["campaignSegmentName"].ToString());
                Console.WriteLine(rows[i]["campaignSegmentName"]);
            }

            return Global.aCampgainSegments = campaginSegment;
        }

        public static List<string> GetHeaderNames(DataTable TheDataTable)
        {

            List<string> sHeaderNames = new List<string>();
            foreach (DataColumn dc in TheDataTable.Columns)
            {

                //Console.WriteLine(dc.ColumnName);
                // Add Column names to array
                sHeaderNames.Add(dc.ColumnName);

            }
            // Print the number of rows in the collection.
            Console.WriteLine("The Number of Rows: " + TheDataTable.Rows.Count);

            return Global.HeaderNames = sHeaderNames;

        }

        

        private static string GetSQLQueryForSSN(string sTheSSN)
        {
            if (Global.isPrimary)
            {
                string queryString = "SELECT * " +
                                "FROM Reporting.V_EligibilityReport " +
                                "Where SocialSecurityNumber = " + sTheSSN;
                Console.WriteLine("SQL: " + queryString);
                return queryString;
            }
            else
            {
                string queryString = "SELECT * " +
                                "FROM Reporting.V_EligibilityReport " +
                                "Where sdpSocialSecurityNumber = " + sTheSSN;

                Console.WriteLine("SQL: " + queryString);
                return queryString;
            }
                      
        }

        private static string GetSQLQueryForCampaignName(List<string> aCampgainSegments)
        {

            StringBuilder sCampaginSegments = new StringBuilder();
         
                string queryString = "SELECT campaignSegmentName, campaignSegmentGuid FROM dbo.CampaignSegment Where campaignSegmentGuid in ('";

                    for (int i = 0; i < aCampgainSegments.Count; i++)
                    {
                        sCampaginSegments.Append(aCampgainSegments[i] + "','");
                    }

                //Trim ending comma
                sCampaginSegments.Length--;
                sCampaginSegments.Length--;

                MessageBox.Show(queryString + sCampaginSegments + ")");
                
                return queryString + sCampaginSegments + ")";
          
           


        }



    }




}