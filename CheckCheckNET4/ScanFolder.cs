using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows;

namespace CheckCheckNET4
{
    class ScanFolder
    {
        public static void StartFileScanning(string conStr)
        {
            string sheetName;
            //Get the name of the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    try
                    {
                        cmd.Connection = con;
                        con.Open();
                        DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                        con.Close();
                    }
                    catch (OleDbException e)
                    {
                        MessageBox.Show("Får ikke åpnet en av Excel-filene. Virker som den er åpen et annet sted...", e.Message);
                        return;
                    }

                }
            }
            ReadDataFromFirstSheet(conStr, sheetName);
        }

        public static void ReadDataFromFirstSheet(string conStr, string sheetName)
        {
            //Read Data from the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();

                        //Populate DataGridView.

                    }
                }
            }

        }
    }
}
