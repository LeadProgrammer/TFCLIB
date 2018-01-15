using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace LIBDI
{
    public class Config
    {
        public string scConfig = "[@TFC_CONFIG]";

        private string SQL = "";

        private SQLDirect cSQL = new SQLDirect();

        private SqlConnection oConn = new SqlConnection();


        public void SetConfig(SAPbobsCOM.Company oComp, string sName, string sVal, string sGrp, string sCmt)
        {
            oConn = cSQL.oConnectToSql(ref oComp);
            SetConfig(oConn, sName, sVal, sGrp, sCmt);
        }

        public void SetConfig(SqlConnection oConn, string sName, string sVal, string sGrp, string sCmt)
        {

            string sCode = "";

            try
            {
                SQL = "SELECT Code FROM " + scConfig + " WITH (NOLOCK) WHERE Name = N'" + sName + "'";
                sCode = cSQL.sGetDBField(ref oConn, SQL);

                // entry does not exist, add it
                if (string.IsNullOrEmpty(sCode))
                {
                    // first find the next code
                    SQL = "SELECT MAX(Code)+10 FROM " + scConfig + " WITH (NOLOCK)";
                    sCode = cSQL.sGetDBField(ref oConn, SQL);
                    if (sCode == "0") sCode = "1000";

                    SQL = "INSERT " + scConfig + " VALUES('" + sCode + "','" + sName + "','" + sVal + "','" + sGrp + "','" + sCmt + "'";


                    if (cSQL.SQLCommand(oConn, SQL) != 1)
                    {
                        //DImsg.MessageERR(iErr, sErr); there is an error message in SQLCommand
                    }
                    return;
                }

                SQL = "";
                SQL = SQL + "UPDATE " + scConfig + " SET U_Value = '" + sVal + "'";
                if (sGrp != null) SQL = SQL + " , U_Group = '" + sGrp + "'";
                if (sCmt != null) SQL = SQL + " , U_Comment = '" + sCmt + "'";
                SQL = SQL + " WHERE Name = '" + sName + "'";

                cSQL.SQLCommand(oConn, SQL);

            }
            catch (Exception ex)
            {
                DImsg.MessageERR(ref ex);
            }
            finally
            {
                GC.Collect();
            }
        }

        public string sGetConfig(SAPbobsCOM.Company oComp, string sName)
        {
            SqlConnection oConn = cSQL.oConnectToSql(ref oComp);

            return sGetConfig(oConn, sName);
        }

        public string sGetConfig(SqlConnection oConn, string sName)
        {

            try
            {
                // get the entry
                string sRec = cSQL.sGetDBField(ref oConn, "SELECT U_Value FROM " + scConfig + " WITH (NOLOCK) WHERE Name = '" + sName + "'");
                return sRec;
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

        public void GetConfigGroup(ref SAPbobsCOM.Company oComp, string sColumn, string sGroup, ref string[] sArray, ref bool BubbleEvent)
        {
            //// routine to load an array from the specified column for all entries that match the group
            //if (BubbleEvent == false)
            //    return;

            //int i = 0;

            //SAPbobsCOM.Recordset oRS1 = null;

            //try
            //{
            //    SQL = "SELECT " + sColumn + " FROM [" + scConfig + "] WITH (NOLOCK) WHERE U_Group = N'" + sGroup + "'";
            //    oRS1 = rsGetDBRecSet(oComp, SQL, BubbleEvent);

            //    if (bLog)
            //        LOG(SQL);
            //    if (bLog)
            //        LOG(oRS1.RecordCount.ToString());

            //    sArray = new string[oRS1.RecordCount];
            //    if (oRS1.RecordCount < 1)
            //        break; // TODO: might not be correct. Was : Exit Try
            //    oRS1.MoveFirst();
            //    for (i = 0; i <= oRS1.RecordCount - 1; i++)
            //    {
            //        sArray[i] = oRS1.Fields.Item(0).Value.ToString();
            //        if (bLog)
            //            LOG(i.ToString() + "   - " + sArray[i]);
            //        oRS1.MoveNext();
            //    }

            //}
            //catch (Exception ex)
            //{
            //    BubbleEvent = false;
            //    MessageERR(Err().Number, ex);
            //}

            //if ((oRS1 != null))
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1);
            //GC.Collect();

        }
        public void GetConfigGroup(ref SAPbobsCOM.Company oComp, string sColumn, string sGroup, ref decimal[] dArray, ref bool BubbleEvent)
        {
            //// routine to load an array from the specified column for all entries that match the group
            //if (BubbleEvent == false)
            //    return;

            //int i = 0;

            //SAPbobsCOM.Recordset oRS1 = null;

            //try
            //{
            //    SQL = "SELECT " + sColumn + " FROM [" + scConfig + "] WITH (NOLOCK) WHERE U_Group = N'" + sGroup + "'";
            //    oRS1 = rsGetDBRecSet(oComp, SQL, BubbleEvent);

            //    dArray = new decimal[oRS1.RecordCount];
            //    if (oRS1.RecordCount < 1)
            //        break; // TODO: might not be correct. Was : Exit Try
            //    oRS1.MoveFirst();
            //    for (i = 0; i <= oRS1.RecordCount - 1; i++)
            //    {
            //        dArray[i] = oRS1.Fields.Item(0).Value;
            //        oRS1.MoveNext();
            //    }

            //}
            //catch (Exception ex)
            //{
            //    BubbleEvent = false;
            //    MessageERR(Err().Number, ex);
            //}

            //if ((oRS1 != null))
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1);
            //GC.Collect();

        }
    }
}
