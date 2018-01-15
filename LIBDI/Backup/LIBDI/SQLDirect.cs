using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace LIBDI
{
    public class SQLDirect
    {
        string sTmp = "";
        public string scSQLName = "SAPAddon";      // also exists in B1Company
        public string scSQLPass = "W!@rM68J^x";    // also exists in B1Company

        public SqlConnection oConnectToSql(ref SAPbobsCOM.Company oCompany)
        {

            int iErr = 0;

            SqlConnection oConn = oConnectToSql(oCompany.Server, oCompany.CompanyDB, scSQLName, scSQLPass, true, ref iErr);

            return oConn;
        }

        public SqlConnection oConnectToSql(string sServer, string sDataBase)
        {
            //if (!bErr) return null;
            // see if the user exists
            //select loginname from master.dbo.syslogins  where name = @loginName and dbname = 'PUBS')

            int iErr = 0;

            SqlConnection oConn = oConnectToSql(sServer, sDataBase, scSQLName, scSQLPass, true, ref iErr);

            // if the user name is invalid then add the SQL user
            if (iErr == 18456)
            {
                iErr = 0;
                CreateSQL_User(sServer, sDataBase, ref iErr);
            }

            if (iErr == 0) oConn = oConnectToSql(sServer, sDataBase, scSQLName, scSQLPass, true, ref iErr);

            return oConn;
        }

        public SqlConnection oConnectToSql(string sServer, string sDataBase, string sUser, string sPassword, bool bDisplayErrMsg, ref int iErr)
        {
            SqlConnection functionReturnValue = default(SqlConnection);

            functionReturnValue = null;

            if (iErr != 0) return functionReturnValue;

            SqlConnection oConn = new SqlConnection();
            string sConnStr = "";
            sConnStr = sConnStr + "integrated security=False;";
            sConnStr = sConnStr + "data source=" + sServer + ";";
            sConnStr = sConnStr + "persist security info=False;";
            sConnStr = sConnStr + "initial catalog=" + sDataBase + ";";
            sConnStr = sConnStr + "User ID=" + sUser + ";";
            sConnStr = sConnStr + "Password=" + sPassword + ";";

            oConn.ConnectionString = sConnStr;

            try
            {
                oConn.Open();
            }
            catch (SqlException ex)
            {
                iErr = ex.Number;

                // don't display error message if the user is addin user is invalid
                if (ex.Number == 18456 & sUser == scSQLName & sPassword == scSQLPass) return (null);  //"Login failed for user '" + sUser + "'."
                if (bDisplayErrMsg)
                    DImsg.MessageERR(ex.Number, ex.Message);          
                return null;
            }
            catch (Exception ex)
            {
                iErr = -1;
                if (bDisplayErrMsg)
                    DImsg.MessageERR(ref ex);
                return null;
            }

            return oConn;
            //return functionReturnValue;

            //oConn.InfoMessage += New SqlInfoMessageEventHandler(OnInfoMessage)
            //oConn.StateChange += New StateChangeEventHandler(OnStateChange)
        }

        public int SQLCommand(SqlConnection oConn, string sCommand)
        {
            int iRtn = -1;
            try
            {
                if (oConn != null)
                {
                    SqlCommand sSQLCmd;
                    using (sSQLCmd = new SqlCommand(sCommand, oConn))
                    {
                        // ExecuteNonQuery returns the number of rows accessed.
                        iRtn = sSQLCmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                iRtn = -1;
                DImsg.MessageERR(ref ex);
            }
            finally
            {
            }
            return (iRtn);
        }

        public object ExecuteScalar(SqlConnection oConn, string sCommand)
        {
            object value = null;
            try
            {
                if (oConn != null)
                {
                    SqlCommand sSQLCmd;
                    using (sSQLCmd = new SqlCommand(sCommand, oConn))
                    {
                        // ExecuteNonQuery returns the number of rows accessed.
                        value = sSQLCmd.ExecuteScalar();
                    }
                }
            }
            catch (Exception ex)
            {
                DImsg.MessageERR(ref ex);
            }
            finally
            {
            }
            return value;
        }

        public DataTable LoadDataTable(SqlConnection oConn, string sQuery, string sTableName)
        {

            if (oConn == null) return null;
            if (sQuery.Trim().Length == 0) return null;
            if (sTableName.Trim().Length == 0) return null;

            try
            {
                DataSet ds = new DataSet();
                SqlDataAdapter da = new SqlDataAdapter(sQuery, oConn);
                ds = new DataSet(sTableName);
                da.Fill(ds, sTableName);
                return ds.Tables[sTableName];
            }
            catch (Exception ex)
            {
                DImsg.MessageERR(ref ex);
                return null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public string sGetDBField(ref SqlConnection oConn, string Query)
        {
            // get the first element - 
            //  if nothing if found return empty.


            string sReturn = "";
            string[][] RecordData;
            if (oConn != null)
            {
                RecordData = ReturnArrayQueryData(oConn, Query);
                if (RecordData == null) return (sReturn);
                if (RecordData.Length <= 0) return (sReturn);
                try
                {
                    sReturn = RecordData[0][0];
                }
                catch (Exception ex)
                {
                    DImsg.MessageERR(ref ex);
                }
            }
            return (sReturn);
        }

        public string[][] ReturnArrayQueryData(SqlConnection oConn, string Query)
        {
            string[][] RecordArray = null;

            if (Query.Length == 0) return null;
            if (oConn == null) return null;

            SqlDataReader oRecordSet = null;

            oRecordSet = ReturnRecordSet(oConn, Query);
            RecordArray = ReturnArrayOfData(oRecordSet);

            if (oRecordSet != null) oRecordSet.Close();

            if (RecordArray.Length != 0)
                return (RecordArray);
            else
                return null; // new string[0][];
        }

        public string[][] ReturnArrayOfData(SqlDataReader oRecordSet)
        {
            ArrayList records = new ArrayList();
            bool bFoundData = false;
            try
            {
                if (oRecordSet == null)
                    bFoundData = false;
                else if (oRecordSet.HasRows == false)
                    bFoundData = false;
                else if (oRecordSet.HasRows == true)
                {
                    try
                    {
                        while (oRecordSet.Read())
                        {
                            bFoundData = true;
                            string[] readerData = new string[oRecordSet.FieldCount];
                            records.Add(readerData);
                            int loopcount = oRecordSet.FieldCount;
                            for (int i = 0; i < loopcount; i++)
                            {
                                try
                                {
                                    readerData[i] = oRecordSet.GetValue(i).ToString();
                                }
                                catch
                                {
                                    readerData[i] = "";
                                }
                            }
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
                bFoundData = false;
            }
            if (bFoundData)
                return records.ToArray(typeof(string[])) as string[][];
            else
                return new string[0][];
        }

        public string[] sArrayOfColumn(SqlConnection oConn, string sQuery, ref int iErr)
        {

            // return an array of values for one Column  

            if (iErr != 0) return null;

            try
            {
                string[][] sData = null;

                sData = ReturnArrayQueryData(oConn, sQuery);

                if (sData == null) return null;

                int j = sData.GetUpperBound(0);

                string[] sRtn = null;
                sRtn = new string[j + 1];

                for (int i = 0; i <= j; i++)
                {
                    sRtn[i] = sData[i][0];
                }

                return sRtn;

            }
            catch (Exception ex)
            {
                iErr = 1;
                DImsg.MessageERR(ex.Message);
                return null;
            }
            finally
            {
            }
        }

        public decimal[] dArrayOfColumn(SqlConnection oConn, string sQuery, ref int iErr)
        {
 
            // return an array of values for one Column  - DECIMAL

            if (iErr != 0) return null;

             try
            {
                string[][] sData = null;

                sData = ReturnArrayQueryData(oConn, sQuery);

                if (sData == null) return null;

                int j = sData[0].Length - 1;

                decimal[] dRtn = null;
                dRtn = new decimal[j + 1];

                for (int i = 0; i <= j; i++)
                {
                    dRtn[i] = Convert.ToDecimal(sData[i][0]);
                }

                return dRtn;

            }
            catch (Exception ex)
            {
                iErr = 1;
                DImsg.MessageERR(ex.Message);
                return null;
            }
            finally
            {
            }
        }

        public string[] sArrayOfRow(SqlConnection oConn, string sQuery, ref int iErr)
        {
            // return an array of returned values for one Row  

            if (iErr != 0) return null;

            try
            {
                string[][] sData = null;

                sData = ReturnArrayQueryData(oConn, sQuery);

                if (sData == null) return null;

                int j = sData[0].Length - 1;

                string[] sRtn = null;
                sRtn = new string[j + 1];

                for (int i = 0; i <= j; i++)
                {
                    sRtn[i] = sData[0][i];
                }

                return sRtn;

            }
            catch (Exception ex)
            {
                iErr = 1;
                DImsg.MessageERR(ex.Message);
                return null;
            }
            finally
            {
            }
        }

        private SqlDataReader ReturnRecordSet(SqlConnection oConn, string sQuery)
        {
            // olny one datareader per connection

            SqlCommand oSqlCommand = null;
            try
            {
                oSqlCommand = new SqlCommand(sQuery, oConn);
                SqlDataReader oRecordSet = oSqlCommand.ExecuteReader();
                return (oRecordSet);
            }
            catch (Exception ex)
            {
                DImsg.MessageERR(ref ex);
                return (null);
            }
        }

        public string sql_fix(string s)
        {
            return s.Replace("'", "''").Replace("\\", "\\\\");
        }

        public void CreateSQL_User(string sServer, string sDatabase, ref int iErr)
         {
             if (iErr != 0) return;
             CreateSQL_User(sServer, sDatabase, "", "", scSQLName, scSQLPass, 5, ref  iErr);
         }

        public void CreateSQL_User(string sServer, string sDatabase, string saUser, string saPassword, 
                                   string AddonUserName, string AddonUserPassword, int connectiontrys, ref int iErr)
         {
             //select count(*) From master.sysxlogins WHERE NAME = 'myUsername' // see if user exists

             if (iErr != 0) return;

             SqlConnection oConn = oConnectToSql(sServer,sDatabase, AddonUserName, AddonUserPassword, false, ref iErr);
             if (oConn != null)
             {                 
                  goto ExitSub;
             }

         tryagain:
             SQLAddUserPrompt oForm = new SQLAddUserPrompt(sServer, sDatabase);
             oForm.GetUserAndPassowrd(ref saUser, ref saPassword);
             if ((oForm.CancelWasPressed == true))
             {
                 goto ExitSub;
             }
             oConn = oConnectToSql(sServer, sDatabase, saUser, saPassword, false, ref iErr);
             if (oConn == null)
             {
                 iErr = 0;
                 goto tryagain;
             }
             //  add the user to the database
             iErr = 1;
             try
             {
                 SqlCommand c = new SqlCommand(("CREATE LOGIN "
                                 + (sql_fix(AddonUserName) + (" WITH PASSWORD = \'"
                                 + (sql_fix(AddonUserPassword) + "\' , CHECK_EXPIRATION = OFF, CHECK_POLICY = OFF ")))), oConn);
                 int inttemp = c.ExecuteNonQuery();
             }
             catch (Exception ex)
             {
                 DImsg.MessageERR("CREATE LOGIN creation failed"  + Environment.NewLine + ex.Message);
                 goto ExitSub;
             }
             try
             {
                 SqlCommand c2 = new SqlCommand(("EXEC sys.sp_addsrvrolemember @loginame ="
                                 + (sql_fix(AddonUserName) + ", @rolename = N\'sysadmin\'")), oConn);
                 int inttemp2 = c2.ExecuteNonQuery();
             }
             catch (Exception ex)
             {
                 DImsg.MessageERR("sys.sp_addsrvrolemember creation failed" + Environment.NewLine + ex.Message);
                 goto ExitSub;
             }
             iErr = 0;

         ExitSub:
             if (oConn != null)
             {
                 oConn.Close();
             }
             GC.Collect();
         }

        public string sGetLastDocEntryAdded(SqlConnection oConn, string sTable, string sBP)
        {
            // get the last docement Entry added for the specified BP.

            string sRtn = "";
            string SQL = "";

            try
            {

                SQL = "SELECT MAX(DocEntry) FROM " + sTable + " WHERE CardCode = '" + sBP + "'";
                sRtn = sGetDBField(ref oConn, SQL);

                return sRtn;
            }
            catch (Exception ex)
            {
                DImsg.MessageERR(ref ex);
                return sRtn;
            }
            finally
            {
                GC.Collect();
            }
        }

        public string sGetLastDocNumAdded(SqlConnection oConn, string sTable, string sBP)
        {
            // get the last docement added for the specified BP.

            string sRtn = "";
            string SQL = "";

            try
            {
                // get the Doc Entry
                sRtn = sGetLastDocEntryAdded(oConn, sTable, sBP);

                // get the corresponding DocNum
                SQL = "SELECT DocNum FROM " + sTable + " WHERE DocEntry = '" + sRtn + "'";
                sRtn = sGetDBField(ref oConn, SQL);

                return sRtn;
            }
            catch (Exception ex)
            {
                DImsg.MessageERR(ref ex);
                return sRtn;
            }
            finally
            {
                GC.Collect();
            }
        }

        public string sGetDocEntry(SqlConnection oConn, string sTable, string sDocNum)
        {
            // get the document ENTRY for the specified DocNum.

            string sRtn = "";
            string SQL = "";

            try
            {
                // get the corresponding DocNum
                SQL = "SELECT DocEntry FROM " + sTable + " WHERE DocNum = '" + sDocNum + "'";
                sRtn = sGetDBField(ref oConn, SQL);

                return sRtn;
            }
            catch (Exception ex)
            {
                DImsg.MessageERR(ref ex);
                return sRtn;
            }
            finally
            {
                GC.Collect();
            }
        }

        public decimal dItemPrice(SqlConnection oConn, string sItemCode, int iListnum, ref int iErr)
        {

            // routine to get an Item price based on the Price List Number

            iErr = -1;
            if (iListnum < 1) return 0.0M;

            try
            {
                string SQL = "";
                SQL = SQL + "SELECT Price FROM ITM1 WHERE ItemCode = '" + sItemCode + "' AND PriceList ='" + iListnum.ToString() + "'";
                sTmp = sGetDBField(ref oConn, SQL);
                if (sTmp == "") return 0.0M;
                iErr = 0;
                return Convert.ToDecimal(sTmp);
            }
            catch (Exception ex)
            {
                DImsg.MessageERR(ref ex);
                return 0.0M;
            }
            finally
            {
                GC.Collect();
            }

        }
    }
}
