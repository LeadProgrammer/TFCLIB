using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace LIBDI
{
    public class B1Company
    {
        string scServer = "DAYTONA";
        string scSQLName = "SAPAddon";      // also exists in SQLDirect
        string scSQLPass = "W!@rM68J^x";    // also exists in SQLDirect
        string scSAPName = "manager";   //"B1i";
        string scSAPPass = "Manag3r";   //"tfc1"; 
        SAPbobsCOM.BoDataServerTypes scServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;

        private SAPbobsCOM.Company oComp = new SAPbobsCOM.Company();

        //public SAPbobsCOM.Company Connect(string sFormTitle, ref int iErr)
        //{
        //    // if the database in the connection is SBO-Common, ask which database to connect to
        //    // otherwise use the connect DB.
        //    return null;
        //}

        public SAPbobsCOM.Company Connect(ref SqlConnection oConn, string sFormTitle, ref int iErr)
        {
            return null;
            //int iErr;

            ////oConnTrg = SQLDirect.oConnectToSql(txbDBServerTrg.Text, cmbSAPDBTrg.Text.Substring(0, cmbSAPDBTrg.Text.IndexOf(" - ")), txbDBUserTrg.Text, txbDBPassTrg.Text, true, bErr);
            //oComp.Server = oConn.DataSource;
            //oComp.DbUserName = oConn.Database;
            //oComp.DbPassword = SQLDirect.scSQLName;
            //oComp.DbServerType = pServerType;   // SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
            //oComp.CompanyDB = pDBName;
            //oComp.UserName = pSAPUser;
            //oComp.Password = pSAPPassword;

            //iErr = oComp.Connect();
            //if (iErr == 0)
            //{
            //    return (oComp);
            //}
            //else
            //{
            //    BubbleEvent = false;
            //    DImsg.MessageERR(iErr, oComp.GetLastErrorDescription());
            //    return (null);
            //}
        }

        public SAPbobsCOM.Company Connect(string pDBName, ref int iErr)
        {
            return Connect(scServer, pDBName, scSQLName, scSQLPass, scServerType, scSAPName, scSAPPass, ref iErr);
                
        }

        public SAPbobsCOM.Company Connect(string pServer, string pDBName, string pDBUser, string pDBPassword, SAPbobsCOM.BoDataServerTypes pServerType,
                                          string pSAPUser, string pSAPPassword, ref int iErr)
        {

            if (pDBName == "") iErr = 1;
       
            if (iErr != 0) return null;

            //oConnTrg = SQLDirect.oConnectToSql(txbDBServerTrg.Text, cmbSAPDBTrg.Text.Substring(0, cmbSAPDBTrg.Text.IndexOf(" - ")), txbDBUserTrg.Text, txbDBPassTrg.Text, true, bErr);
            oComp.Server = pServer;
            if (pServer == "") oComp.Server = scServer;

            oComp.DbUserName = pDBUser;
            if (pDBUser == "") oComp.DbUserName = scSQLName;

            oComp.DbPassword = pDBPassword;
            if (pDBPassword == "") oComp.DbPassword = scSQLPass;

            oComp.DbServerType = pServerType;
            //if (pServerType == null) oComp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;

            oComp.CompanyDB = pDBName;

            oComp.UserName = pSAPUser;
            if (pSAPUser == "") oComp.UserName = scSAPName;

            oComp.Password = pSAPPassword;
            if (pSAPPassword == "") oComp.Password = scSAPPass;

            iErr = oComp.Connect();
            if (iErr == 0)
            {
                return oComp;
            }
            else
            {
                DImsg.MessageERR(iErr, oComp.GetLastErrorDescription());
                return null;
            }
        }

    }
}
