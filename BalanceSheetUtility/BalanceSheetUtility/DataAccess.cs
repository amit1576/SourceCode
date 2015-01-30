using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace BalanceSheetUtility
{
    public class DataAccess 
    {
        DataSet t_ReturnDataSet = new DataSet();

        BalanceSheetUtility t_MainClass = new BalanceSheetUtility();
   
        public int ExecuteDataAccessStatement(string p_SQLQuery, string p_ConnectionString)
        {               
            t_ReturnDataSet = new DataSet();
            string t_con = p_ConnectionString;

            using (IDbConnection t_connection = new OleDbConnection(t_con))
            {
                using (IDbCommand t_dbCommand = new OleDbCommand())
                {
                    try
                    {
                        IDbDataAdapter t_dataAdapter = new OleDbDataAdapter();

                        t_dbCommand.CommandText = p_SQLQuery;
                        t_dbCommand.Connection = t_connection;
                        t_dataAdapter.SelectCommand = t_dbCommand;
                        t_dataAdapter.Fill(t_ReturnDataSet);
                        t_dataAdapter = null;
                    }
                    catch (OleDbException t_exception)
                    {
                        t_MainClass.LogProgress("Error While Fetching Data from Source" + t_exception.Message);
                        //t_MainClass.ProfileLogic(t_exception.Message.ToString());
                        //t_MainClass.ProfileLogic(t_exception.StackTrace);
                        return 1;
                    }
                }
            }

            return 0;

        }

        public int ExcuteStoredProcedure(string p_SpStatement, string p_ConnectionString)
        {
            string t_con = p_ConnectionString;

            using (IDbConnection t_connection = new OleDbConnection(t_con))
            {
                using (IDbCommand t_dbCommand = new OleDbCommand())
                {
                    try
                    {
                        IDbDataAdapter t_dataAdapter = new OleDbDataAdapter();

                        t_dbCommand.CommandText = p_SpStatement;
                        t_dbCommand.Connection = t_connection;
                        t_dataAdapter.SelectCommand = t_dbCommand;
                        t_dataAdapter.Fill(t_ReturnDataSet);
                        t_dataAdapter = null;
                    }
                    catch (OleDbException t_exception)
                    {
                        t_MainClass.LogProgress("Error While Executing Stored Procedure : " + t_exception.Message);
                        //t_MainClass.ProfileLogic(t_exception.Message.ToString());
                        //t_MainClass.ProfileLogic(t_exception.StackTrace);
                        return 1;
                    }
                }
            }

            return 0;
        }

        public DataSet ReturnDataSet { get { return t_ReturnDataSet; } }


    }
}