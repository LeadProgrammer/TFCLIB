using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LIBDI
{
    class B1SQL
    {
        string SQL = "";
        int iTmp = 0;

        public SAPbobsCOM.Recordset rsGetDBRecSet(ref SAPbobsCOM.Company oComp, string sSQL, ref bool BubbleEvent)
        {
            SAPbobsCOM.Recordset functionReturnValue = default(SAPbobsCOM.Recordset);

            // routine to execut the sql command and return a record set 

            functionReturnValue = null;
            if (BubbleEvent == false)
                return functionReturnValue;

            //Dim oRS1 As SAPbobsCOM.Recordset = Nothing

            try
            {

                //oRS1 = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                //oRS1.DoQuery(sSQL)
                //rsGetDBRecSet = oRS1
                functionReturnValue = (SAPbobsCOM.Recordset)oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                functionReturnValue.DoQuery(sSQL);
                iTmp = functionReturnValue.RecordCount;
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                DImsg.MessageERR(ref ex);
            }

            //If Not oRS1 Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1)
            GC.Collect();
            return functionReturnValue;

        }
        public string sGetDBfield(ref SAPbobsCOM.Company oComp, string sSQL, ref bool BubbleEvent)
        {
            string functionReturnValue = null;

            // routine to execut the sql command and return a value 

            functionReturnValue = "";
            if (BubbleEvent == false)
                return functionReturnValue;

            SAPbobsCOM.Recordset oRS1 = null;

            try
            {
                oRS1 = rsGetDBRecSet(ref oComp, sSQL, ref BubbleEvent);
                if (oRS1.RecordCount > 0)
                    functionReturnValue = oRS1.Fields.Item(0).Value.ToString();

            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                DImsg.MessageERR(ref ex);
            }

            if ((oRS1 != null))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1);
            GC.Collect();
            return functionReturnValue;

        }
        public string sGetDBfield(ref SAPbobsCOM.Company oComp, string sSQL, ref int iRecCnt, ref bool BubbleEvent)
        {
            string functionReturnValue = null;

            // routine to execut the sql command and return a value 

            functionReturnValue = "";
            if (BubbleEvent == false)
                return functionReturnValue;

            SAPbobsCOM.Recordset oRS1 = null;

            try
            {
                oRS1 = rsGetDBRecSet(ref oComp, sSQL, ref BubbleEvent);
                iRecCnt = oRS1.RecordCount;
                if (oRS1.RecordCount > 0)
                    functionReturnValue = oRS1.Fields.Item(0).Value.ToString();

            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                DImsg.MessageERR(ref ex);
            }

            if ((oRS1 != null))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1);
            GC.Collect();
            return functionReturnValue;

        }
        public int iGetDBfield(ref SAPbobsCOM.Company oComp, string sSQL, ref bool BubbleEvent)
        {
            int functionReturnValue = 0;

            // routine to execut the sql command and return a value 

            functionReturnValue = 0;
            if (BubbleEvent == false)
                return functionReturnValue;

            SAPbobsCOM.Recordset oRS1 = null;

            try
            {
                oRS1 = rsGetDBRecSet(ref oComp, sSQL, ref BubbleEvent);
                if (oRS1.RecordCount > 0)
                    functionReturnValue = Convert.ToInt32(oRS1.Fields.Item(0).Value);

            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                DImsg.MessageERR(ref ex);
            }

            if ((oRS1 != null))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1);
            GC.Collect();
            return functionReturnValue;

        }
        public int iGetDBfield(ref SAPbobsCOM.Company oComp, string sSQL, ref int iRecCnt, ref bool BubbleEvent)
        {
            int functionReturnValue = 0;

            // routine to execut the sql command and return a value 

            functionReturnValue = 0;
            if (BubbleEvent == false)
                return functionReturnValue;

            SAPbobsCOM.Recordset oRS1 = null;

            try
            {
                oRS1 = rsGetDBRecSet(ref oComp, sSQL, ref BubbleEvent);
                iRecCnt = oRS1.RecordCount;
                if (oRS1.RecordCount > 0)
                    functionReturnValue = Convert.ToInt32(oRS1.Fields.Item(0).Value);

            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                DImsg.MessageERR(ref ex);
            }

            if ((oRS1 != null))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1);
            GC.Collect();
            return functionReturnValue;

        }
        public decimal dGetDBfield(ref SAPbobsCOM.Company oComp, string sSQL, ref bool BubbleEvent)
        {
            decimal functionReturnValue = default(decimal);

            // routine to execut the sql command and return a value 

            functionReturnValue = 0m;
            if (BubbleEvent == false)
                return functionReturnValue;

            SAPbobsCOM.Recordset oRS1 = null;

            try
            {
                oRS1 = rsGetDBRecSet(ref oComp, sSQL, ref BubbleEvent);
                if (oRS1.RecordCount > 0)
                    functionReturnValue = Convert.ToDecimal(oRS1.Fields.Item(0).Value);

            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                DImsg.MessageERR(ref ex);
            }

            if ((oRS1 != null))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1);
            GC.Collect();
            return functionReturnValue;

        }
        public decimal dGetDBfield(ref SAPbobsCOM.Company oComp, string sSQL, ref int iRecCnt, ref bool BubbleEvent)
        {
            decimal functionReturnValue = default(decimal);

            // routine to execut the sql command and return a value 

            functionReturnValue = 0m;
            if (BubbleEvent == false)
                return functionReturnValue;

            SAPbobsCOM.Recordset oRS1 = null;

            try
            {
                oRS1 = rsGetDBRecSet(ref oComp, sSQL, ref BubbleEvent);
                iRecCnt = oRS1.RecordCount;
                if (oRS1.RecordCount > 0)
                    functionReturnValue = Convert.ToDecimal(oRS1.Fields.Item(0).Value);

            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                DImsg.MessageERR(ref ex);
            }

            if (oRS1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1);
            GC.Collect();
            return functionReturnValue;

        }
        public string sGetDocNumF(string sDocEntry, string sFormNum, ref bool BubbleEvent)
        {
            string functionReturnValue = null;
            //// get the docnum using the form number (139, 149, ect)
            //functionReturnValue = "";
            //if (BubbleEvent == false)
            //    return functionReturnValue;
            //try
            //{
            //    SQL = "SELECT DocNum FROM " + sGetMDTableF(sFormNum, "T") + " WHERE DocEntry = '" + sDocEntry + "'";
            //    functionReturnValue = sGetDBfield(ref oComp, System.Data.SQL, ref BubbleEvent);
            //}
            //catch (Exception ex)
            //{
            //    BubbleEvent = false;
            //    DImsg.MessageERR(ref ex);
            //}
            //GC.Collect();
            return functionReturnValue;
        }
        public string sGetDocNumT(string sDocEntry, string sTable, ref bool BubbleEvent)
        {
            string functionReturnValue = null;
            //// get the docnum using the Table
            //functionReturnValue = "";
            //if (BubbleEvent == false)
            //    return functionReturnValue;
            //try
            //{
            //    System.Data.SQL = "SELECT DocNum FROM " + sTable + " WHERE DocEntry = '" + sDocEntry + "'";
            //    functionReturnValue = sGetDBfield(ref oComp, System.Data.SQL, ref BubbleEvent);
            //}
            //catch (Exception ex)
            //{
            //    BubbleEvent = false;
            //    DImsg.MessageERR(ref ex);
            //}
            GC.Collect();
            return functionReturnValue;
        }
        public string sNextNum(bool bUpdate, ref bool BubbleEvent)
        {
            string functionReturnValue = null;

            // this routine looks at the ONNM and NNM1 table using the B1 objects.

            // this routine is not functioning

            // There is no way to update it at this time

            //functionReturnValue = "";
            //if (BubbleEvent == false)
            //    return functionReturnValue;

            //SAPbobsCOM.CompanyService oCmpSrv = default(SAPbobsCOM.CompanyService);
            //SAPbobsCOM.SeriesService oSeriesService = default(SAPbobsCOM.SeriesService);
            //SAPbobsCOM.Series oSeries = default(SAPbobsCOM.Series);
            //SAPbobsCOM.SeriesParams oSeriesParams = default(SAPbobsCOM.SeriesParams);

            ////get company service
            //oCmpSrv = oComp.GetCompanyService;

            ////get series service
            //oSeriesService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService);

            ////get series
            //oSeries = oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiSeries);

            ////set series name
            //oSeries.Name = "Series1";

            ////set doument type(e.g. Deliveries=15)
            //oSeries.Document = 15;

            ////set the period indicator
            //oSeries.PeriodIndicator = "Default";

            ////set the group code
            ////(enum BoSeriesGroupEnum has all Group Enum)
            //oSeries.GroupCode = 1;

            ////set the first number
            //oSeries.InitialNumber = 300;
            ////set last number
            //oSeries.LastNumber = 350;

            ////add series
            ////before adding the series to the document check that the lastNumber property
            ////of the last series has a value(if not the add function will fail)
            //oSeriesParams = oSeriesService.AddSeries(oSeries);
            return functionReturnValue;
        }
        public string sNextNum(string sTable, string sField, string sPrefix, int iBase, ref bool BubbleEvent)
        {
            string functionReturnValue = null;

            //// routine to get the next number from the specified field and table.
            //// find the largest number and add 1

            //// sTable    - the table
            //// sField    - the field
            //// sPrefix   - ie C for customer
            //// iBase     - this is in case there is no entry already

            //// the increment is 1

            //functionReturnValue = "";
            //if (BubbleEvent == false)
            //    return functionReturnValue;



            //try
            //{
            //    functionReturnValue = sNextNum(sTable, sField, sPrefix, iBase, 1, ref BubbleEvent);

            //}
            //catch (Exception ex)
            //{
            //    BubbleEvent = false;
            //    DImsg.MessageERR(ref ex);
            //}
            //finally
            //{
            //    // exit try ends up here, but the objects must be instantiated otherwise there will be and error.
            //    //If Not oRS Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
            //    GC.Collect();
            //}
            return functionReturnValue;

        }
        public string sNextNum(string sTable, string sField, string sPrefix, int iBase, int iIncr, ref bool BubbleEvent)
        {
            string functionReturnValue = null;

            // routine to get the next number from the specified field and table.
            //// find the largest number and add 1

            //// sTable    - the table
            //// sField    - the field
            //// sPrefix   - ie C for customer
            //// iBase     - this is in case there is no entry already
            //// iIncr      - the increment, like 1 or 10

            //functionReturnValue = "";
            //if (BubbleEvent == false)
            //    return functionReturnValue;

            //int iPreLen = sPrefix.Length;
            //// length of the prefix

            //int iLen = 20;

            //SAPbobsCOM.Recordset oRS = null;


            //try
            //{
            //    oRS = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //    System.Data.SQL = "";
            //    System.Data.SQL = "CAST(" + sField + " AS VARCHAR(" + iLen.ToString() + "))";
            //    System.Data.SQL = "ISNULL(" + System.Data.SQL + ", '" + sPrefix + "' + CAST(" + iBase.ToString() + " AS VARCHAR(" + iLen.ToString() + ")))";
            //    iPreLen = iPreLen + 1;
            //    System.Data.SQL = "SUBSTRING (" + System.Data.SQL + ", " + iPreLen.ToString() + ", " + iLen.ToString() + ")";
            //    System.Data.SQL = "SELECT MAX(CAST(" + System.Data.SQL + " AS INTEGER))+" + iIncr.ToString() + " FROM [" + sTable + "] WITH (NOLOCK) ";
            //    System.Data.SQL = System.Data.SQL + " WHERE " + sField + " LIKE '" + sPrefix + "%'";

            //    oRS.DoQuery(System.Data.SQL);
            //    functionReturnValue = oRS.Fields.Item(0).Value.ToString;
            //    if (functionReturnValue == null | string.IsNullOrEmpty(functionReturnValue) | functionReturnValue == "0")
            //        functionReturnValue = Convert.ToString(iBase);
            //    functionReturnValue = sPrefix + functionReturnValue;

            //}
            //catch (Exception ex)
            //{
            //    BubbleEvent = false;
            //    DImsg.MessageERR(ref ex);
            //}
            //finally
            //{
            //    // exit try ends up here, but the objects must be instantiated otherwise there will be and error.
            //    if ((oRS != null))
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            //    GC.Collect();
            //}
            return functionReturnValue;

        }
        public string sGetAuthorization(string sID, ref bool BubbleEvent)
        {
            string functionReturnValue = null;

            // function to return the authorization for an ID

            functionReturnValue = "";

            //if (BubbleEvent == false)
            //    return functionReturnValue;

            //string SQL = null;
            //SAPbobsCOM.Recordset oRS = null;

            ////Dim oUPT As SAPbobsCOM.UserPermissionTree = Nothing
            ////oUPT = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
            ////If oUPT.GetByKey(sID) = False Then Exit Try
            ////sGetAuthorization = oUPT.Options()


            //try
            //{
            //    oRS = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //    SQL = "";
            //    SQL = SQL + "SELECT T1.SUPERUSER, T0.Permission ";
            //    SQL = SQL + "  FROM OUSR T1 WITH (NOLOCK)";
            //    SQL = SQL + "  LEFT OUTER JOIN USR3 T0 ";
            //    SQL = SQL + "    ON T0.UserLink = T1.UserId ";
            //    SQL = SQL + "   AND T0.PermID = '" + sID + "'";
            //    SQL = SQL + " WHERE T1.USER_CODE = '" + oComp.UserName + "' ";
            //    oRS.DoQuery(SQL);

            //    if (oRS.RecordCount < 1)
            //        break; // TODO: might not be correct. Was : Exit Try

            //    functionReturnValue = oRS.Fields.Item("Permission").Value;

            //    if (oRS.Fields.Item("SUPERUSER").Value == "Y")
            //        functionReturnValue = "F";

            //}
            //catch (Exception ex)
            //{
            //    BubbleEvent = false;
            //    DImsg.MessageERR(ref ex);
            //}
            //finally
            //{
            //    // exit try ends up here, but the objects must be instantiated otherwise there will be and error.
            //    if ((oRS != null))
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            //    GC.Collect();
            //}
            return functionReturnValue;

        }
        public string sCurrentEmp(string sField, ref bool BubbleEvent)
        {
            string functionReturnValue = null;

            // function to return the specified field from OHEM using the logged on user

            functionReturnValue = "";

            //if (BubbleEvent == false)
            //    return functionReturnValue;

            //string SQL = null;
            //SAPbobsCOM.Recordset oRS = null;


            //try
            //{
            //    oRS = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //    SQL = "";
            //    SQL = SQL + "SELECT * FROM OHEM WHERE userID = '" + oComp.UserSignature.ToString + "' ";
            //    oRS.DoQuery(SQL);

            //    if (oRS.RecordCount < 1)
            //        break; // TODO: might not be correct. Was : Exit Try

            //    functionReturnValue = oRS.Fields.Item(sField).Value;

            //}
            //catch (Exception ex)
            //{
            //    BubbleEvent = false;
            //    DImsg.MessageERR(ref ex);
            //}
            //finally
            //{
            //    // exit try ends up here, but the objects must be instantiated otherwise there will be and error.
            //    if ((oRS != null))
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
            //    GC.Collect();
            //}
            return functionReturnValue;

        }
    }
}
