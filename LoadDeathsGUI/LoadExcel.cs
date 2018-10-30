using System;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Globalization;
using System.Windows.Forms;
using LoadDeathsGUI;
using System.Collections.Generic;

public static class CreateExcelWorksheet
{


    public static void OpenExcel(System.Data.DataTable TheDataTable)
    {
        //Start Excel
        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

        if (xlApp == null)
        {
            MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.", "DOD Loader", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
            Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
            System.Windows.Forms.Application.Exit();
        }
        //Determines whether the Excel is visible.
        xlApp.Visible = false;
        xlApp.ScreenUpdating = false;
        xlApp.DisplayAlerts = false;


        //Select a range of cells and Set the format of the cells as text.
        GetCellRange(xlApp, "A1:A1", "CD1");
        //Add The Values from the DataTable into the Excel sheet.        
        PrintValues(TheDataTable, GetExcelWorkSheet(xlApp));

        //Make the Window Visable 
        xlApp.Visible = true;
        xlApp.ScreenUpdating = true;
        xlApp.DisplayAlerts = true;

    }

    public static string DateTimeCheck(string dateString)
    {
        DateTime dt = DateTime.ParseExact(dateString, "MMddyyyy", CultureInfo.InvariantCulture);
        return dt.ToString("yyyyMMdd");
        
    }

    private static Range GetCellRange(Microsoft.Office.Interop.Excel.Application xlApp, string StartingRow, string EndingCell)
    {
        // Select the Excel cells, in the range A1 to CH1 in the worksheet.
        Range aRange = GetExcelWorkSheet(xlApp).get_Range(StartingRow, EndingCell);

        if (aRange == null)
        {
            MessageBox.Show("Could not get a range. Check to be sure you have the correct versions of the office DLLs.", "Death\\Divorce Loader", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
            Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            System.Windows.Forms.Application.Exit();
        }



        return aRange;
    }

    public static Worksheet GetExcelWorkSheet(Microsoft.Office.Interop.Excel.Application xlApp)
    {
        //Creates a new workbook. The new workbook becomes the active workbook. Returns a Workbook object.
        Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

        //Returns a Sheets collection that represents all the worksheets in the specified workbook. Read-only Sheets object.
        Worksheet ws = (Worksheet)wb.Worksheets[1];
        if (ws == null)
        {
            MessageBox.Show("Worksheet could not be created. Check that your office installation and project references are correct.", "Death\\Divorce Loader", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
            Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            System.Windows.Forms.Application.Exit();
        }

        //Return WorkSheet
        return ws;

    }

    public static void SetEventCode(Worksheet ws, string sTheValue)
    {
        //This is the eventcode column P
        //Should be 02 for a death and 03 for a divorce
        Range eventCode = ws.Rows.Cells[2, 16];
        Range eligibilityEndDate = ws.Rows.Cells[2, 61];


        //Find the Dependent eligibility and event colmun numbers 
        Range sdpEventCode = ws.Rows.Cells[2, 32];
        Range sdpEligibilityEndDate = ws.Rows.Cells[2, 62];

        //Check for nothing
        if (eventCode.Value == null)
        {
            //MessageBox.Show("Primary Event Code is Empty");
        }

        if (CheckDeathDateExists(eventCode, eligibilityEndDate, sdpEventCode, sdpEligibilityEndDate))
        {
            return;
        }


        //Print the value for the primary or dependant.  
        if (LoadDeathsGUI.Global.isPrimary)
        {
            //Set The  color and Value of the cell for the Primary
            eventCode.Interior.Color = 255;
            eventCode.Value = sTheValue;
            eligibilityEndDate.Interior.Color = 255;
            eligibilityEndDate.Value = LoadDeathsGUI.Global.sTheSelectedDate;

        }
        else
        {
            //MessageBox.Show("dependant");
            //Set The  color and Value of the cell for the Dependent
            sdpEventCode.Interior.Color = 255;
            sdpEventCode.Value = sTheValue;
            sdpEligibilityEndDate.Interior.Color = 255;
            sdpEligibilityEndDate.Value = LoadDeathsGUI.Global.sTheSelectedDate;
        }

    }

    public static bool CheckDeathDateExists(Range eventCode, Range eligibilityEndDate, Range sdpEventCode, Range sdpEligibilityEndDate)
    {

        if (LoadDeathsGUI.Global.isPrimary)
        {
            if (eventCode.Value == "02")
            {
                //Change Cell Color
                //ask user if they want to override the allready loaded data

                eventCode.Interior.Color = 255;
                eligibilityEndDate.Interior.Color = 255;
                DialogResult sDR = MessageBox.Show("Would you like to replace the primarys existing event code and end date?", "Death Already Loaded", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (sDR == DialogResult.Yes)
                {
                    //Replace The old data with the new
                    return false;
                }
                else if (sDR == DialogResult.No)
                {
                    return true;
                }
                
            }
            else if (eventCode.Value == "03")
            {
                eventCode.Interior.Color = 255;
                eligibilityEndDate.Interior.Color = 255;
                DialogResult sDR = MessageBox.Show("Would you like to replace the primarys existing event code and end date?", "Divorce Already Loaded", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (sDR == DialogResult.Yes)
                {
                    //Replace The old data with the new
                    return false;
                }
                else if (sDR == DialogResult.No)
                {
                    return true;
                }
            }
            else if (eligibilityEndDate.Value != null)
            {
                //in the rare case where only the eligiblity end date exists with no event code
                eventCode.Interior.Color = 255;
                eligibilityEndDate.Interior.Color = 255;
                DialogResult sDR = MessageBox.Show("Would you like to replace the primarys existing event code and end date?", "Only End Date is Loaded", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (sDR == DialogResult.Yes)
                {
                    //Replace The old data with the new
                    return false;
                }
                else if (sDR == DialogResult.No)
                {
                    return true;
                }
            }
        }
        else
        {
            //ask user if they want to override the allready loaded data
            if (sdpEventCode.Value == "02")
            {
                //Change Cell Color
                sdpEventCode.Interior.Color = 255;
                sdpEligibilityEndDate.Interior.Color = 255;
                DialogResult sDR = MessageBox.Show("Would you like to replace the dependents existing event code and end date?", "Death Already Loaded", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (sDR == DialogResult.Yes)
                {
                    //Replace The old data with the new
                    return false;
                }
                else if (sDR == DialogResult.No)
                {
                    return true;
                }
            }
            else if (sdpEventCode.Value == "03")
            {
                sdpEventCode.Interior.Color = 255;
                sdpEligibilityEndDate.Interior.Color = 255;
                DialogResult sDR = MessageBox.Show("Would you like to replace the dependents existing event code and end date?", "Divorce Already Loaded", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (sDR == DialogResult.Yes)
                {
                    //Replace The old data with the new
                    return false;
                }
                else if (sDR == DialogResult.No)
                {
                    return true;
                } 

            }
            else if (sdpEligibilityEndDate.Value != null)
            {
                //in the rare case where only the eligiblity end date exists with no event code
                sdpEventCode.Interior.Color = 255;
                sdpEligibilityEndDate.Interior.Color = 255;
                DialogResult sDR = MessageBox.Show("Would you like to replace the dependents existing event code and end date?", "Only End Date is Loaded", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (sDR == DialogResult.Yes)
                {
                    //Replace The old data with the new
                    return false;
                }
                else if (sDR == DialogResult.No)
                {
                    return true;
                }
            }
        }

        return false;

    }

    public static void PrintValues(System.Data.DataTable table, Worksheet ws)
    {

        //Take Array of Headers add them to the first row
        List<string> sHeaderNames = new List<string>();

        int x = 1;
        foreach (string Header in LoadDeathsGUI.Global.HeaderNames)
        {
            Range row1 = ws.Rows.Cells[1, x];
            row1.NumberFormat = "@";
            row1.Value = Header;
            x++;
        }
      
        int y = 2;
        foreach (DataRow row in table.Rows)
        {
            int i = 1;
            foreach (DataColumn column in table.Columns)
            {
                Range row2 = ws.Rows.Cells[y, i];
                Range eventCode = ws.Rows.Cells[y, 16];


                // We check System.Guid becuase excel does not have a native varable type matching GUID so we convert it to a string.
                // To avoid brackets being placed in the empty guid we check if row[column] is not empty
                if (column.DataType.ToString() == "System.Guid" && row[column].ToString() != string.Empty)
                {
                    row2.NumberFormat = "@";
                    string sGUIDBrack = "{" + row[column].ToString() + "}";
                    row2.Value = sGUIDBrack;

                    //Console.WriteLine(row[column].ToString());
                }
                else
                {
                    row2.NumberFormat = "@";
                    row2.Value = row[column];

                    //Console.WriteLine(row[column]);

                }

                i++;

                //for some reason there are always two rows kicking out of loop 
                if (i > 82)
                { break; }

            }

            y++;
        }
        
        if (LoadDeathsGUI.Global.isDeath)
        {
            SetEventCode(ws, "02");
            return;
        }
        else if (LoadDeathsGUI.Global.isDivorce)
        {
            //Check that the primary is not divorcing them self.
            if (LoadDeathsGUI.Global.isPrimary)
            {
                MessageBox.Show("The primary cannot divorce them self."); 
                return;
            }
            else
            {
                SetEventCode(ws, "03");
                return;
            }
            
        }

    }


}