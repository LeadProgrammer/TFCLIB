using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using LIBDI;
namespace LIBUI
{
    public class LIBUI
    {
        private string sTmp = "";
        private string SQL = "";
        private int iTmp;
        private static string scConfig = "";
        private static string scVersion = "";
        private static string scNameSpace = "";
        public const string scUIDI = "UI";
        private LIBDI.LIBDI oLIBDI;
        private LIBDI.MicroSoftWindows oLIBDIWIN;
        private static SAPbobsCOM.Company oComp;
        private static SAPbouiCOM.Application oApp;

        public LIBUI()
        {
        }

        public void SetCompApp(ref SAPbobsCOM.Company oCompany, ref SAPbouiCOM.Application oApplication, string sConfig, string sVersion, string sNameSpace)
        {
            // set the company and application objects
            // this should be done only once or when the compnay changes.

            scConfig = sConfig;
            scVersion = sVersion;
            scNameSpace = sNameSpace;

            oApp = oApplication;
            oComp = oCompany;
            //oLIBDI = new LIBDI();
            //////////////////////////////oLIBDI.SetCompany(oComp, scConfig, scVersion, scUIDI, scNameSpace);

            //Dim oC As New SAPbobsCOM.Company
            //oC = oApp.Company.GetDICompany()    ' <<<<<<<<<<<<<<<<<<<< use this  - don't, this is slow.

        }
        #region "Form Create"
        public int iGridRowLevel(ref SAPbouiCOM.Application oApp, ref SAPbouiCOM.Grid ogrid, ref int iErr)
        {
            // function to return the tree level of the selected row
            // 1 is the top level

            int i1 = 0;

            int iRow = 0;

            if (iErr != 0) return 0;


            try
            {
                iRow = ogrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                while (iRow >= 0)
                {
                    i1 = i1 + 1;
                    iRow = ogrid.Rows.GetParent(iRow);
                }

                return i1;

            }
            catch (Exception ex)
            {
                iErr = 1;
                LIBDI.DImsg.MessageERR(ref ex);
                return 0;
            }
            finally
            {
                // exit try ends up here, but the objects must be instantiated otherwise there will be and error.
                //If Not oForm Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                GC.Collect();
            }
        }

        public void ComboBoxAddVals(ref SAPbouiCOM.Form oForm, string sComboBox, string sSQL, ref int iErr)
        {

            // routine to add values to a combo box based on a query.

            // IF THERE IS FATAL ERROR ON THIS STATEMENT         : oGrid.Columns.Item("Territory").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            // THE SOULUTION IS TO CASTR THE RESULT AS CHARACTER : CAST([@ABY_COMMGROUPS].U_territry AS VARCHAR (20))

            int i1 = 0;
            int iCnt = 0;

            SAPbobsCOM.Recordset oRS = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.ComboBox oCmbBox = default(SAPbouiCOM.ComboBox);
            SAPbouiCOM.ComboBoxColumn oCmbBoxC = default(SAPbouiCOM.ComboBoxColumn);

            if (iErr != 0)
                return;

            try
            {
                oRS = (SAPbobsCOM.Recordset)oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oRS.DoQuery(sSQL);
                if (oRS.RecordCount < 1)  return;

                oItem = oForm.Items.Item(sComboBox);
                oItem.DisplayDesc = true;

                switch (oItem.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        oCmbBox = (SAPbouiCOM.ComboBox)oItem.Specific;

                        if (sSQL.Length == 0)
                        {
                            // just remove the valid values
                            iCnt = oCmbBox.ValidValues.Count;
                            if (iCnt < 1)
                               return;

                            for (i1 = iCnt - 1; i1 >= 0; i1 += -1)
                            {
                                oCmbBox.ValidValues.Remove(i1, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                           return;
                        }

                        for (i1 = 1; i1 <= oRS.RecordCount; i1++)
                        {
                            try
                            {
                                oCmbBox.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), oRS.Fields.Item(1).Value.ToString());
                            }
                            catch (Exception ex)
                            {
                                sTmp = "Entry already exists: " + oRS.Fields.Item(0).Value.ToString() + " - " + oRS.Fields.Item(1).Value.ToString() + Environment.NewLine;
                                sTmp = sTmp + sSQL + Environment.NewLine + "(ComboBoxAddVals)";
                                oApp.MessageBox(sTmp, 1, "Ok", "", "");
                            }
                            oRS.MoveNext();
                        }


                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                        oCmbBoxC = (SAPbouiCOM.ComboBoxColumn)oItem.Specific;

                        if (sSQL.Length == 0)
                        {
                            // just remove the valid values
                            iCnt = oCmbBoxC.ValidValues.Count;
                            if (iCnt < 1)
                               return;

                            for (i1 = iCnt - 1; i1 >= 0; i1 += -1)
                            {
                                oCmbBoxC.ValidValues.Remove(i1, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                           return;
                        }

                        for (i1 = 1; i1 <= oRS.RecordCount; i1++)
                        {
                            try
                            {
                                oCmbBoxC.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), oRS.Fields.Item(1).Value.ToString());
                            }
                            catch (Exception ex)
                            {
                                sTmp = "Entry already exists: " + oRS.Fields.Item(0).Value.ToString()+ " - " + oRS.Fields.Item(1).Value.ToString()+ Environment.NewLine;
                                sTmp = sTmp + sSQL;
                                oApp.MessageBox(sTmp + Environment.NewLine + "(ComboBoxAddVals)", 1, "Ok", "", "");
                            }
                            oRS.MoveNext();
                        }

                        break;
                    default:
                       return;

                        break;
                }

            }
            catch (Exception ex)
            {
                iErr = 1;
                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                // exit try ends up here, but the objects must be instantiated otherwise there will be and error.
                if (oRS != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                GC.Collect();
            }

        }

        public void ComboBoxAddVals(ref SAPbouiCOM.ComboBox oComboBox, string sSQL, ref int iErr)
        {
            // routine to add values to a combo box based on a query.

            // IF THERE IS FATAL ERROR ON THIS STATEMENT         : oGrid.Columns.Item("Territory").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            // THE SOULUTION IS TO CAST THE RESULT AS CHARACTER : CAST([@ABY_COMMGROUPS].U_territry AS VARCHAR (20))

            if (iErr != 0)
                return;

            int i1 = 0;
            int iCnt = 0;

            SAPbobsCOM.Recordset oRS = null;

            try
            {
                oRS = (SAPbobsCOM.Recordset)oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (sSQL.Length == 0)
                {
                    // just remove the valid values
                    iCnt = oComboBox.ValidValues.Count;
                    if (iCnt < 1)
                        return;

                    for (i1 = iCnt - 1; i1 >= 0; i1 += -1)
                    {
                        try
                        {
                            oComboBox.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), oRS.Fields.Item(1).Value.ToString());
                        }
                        catch (Exception ex)
                        {
                            sTmp = "Entry already exists: " + oRS.Fields.Item(0).Value.ToString()+ " - " + oRS.Fields.Item(1).Value.ToString()+ Environment.NewLine;
                            sTmp = sTmp + sSQL;
                            oApp.MessageBox(sTmp + Environment.NewLine + "(ComboBoxAddVals)", 1, "Ok", "", "");
                        }
                    }
                   return;
                }

                oRS.DoQuery(sSQL);
                if (oRS.RecordCount < 1)
                   return;

                for (i1 = 1; i1 <= oRS.RecordCount; i1++)
                {
                    oComboBox.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), oRS.Fields.Item(1).Value.ToString());
                    oRS.MoveNext();
                }

            }
            catch (Exception ex)
            {
                iErr = 1;
                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                // exit try ends up here, but the objects must be instantiated otherwise there will be and error.
                if (oRS != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                GC.Collect();
            }

        }

        public void ComboBoxAddVals(ref SAPbouiCOM.ComboBoxColumn oComboBox, string sSQL, ref int iErr)
        {
            // routine to add values to a combo box based on a query.

            // IF THERE IS FATAL ERROR ON THIS STATEMENT         : oGrid.Columns.Item("Territory").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            // THE SOULUTION IS TO CASTR THE RESULT AS CHARACTER : CAST([@ABY_COMMGROUPS].U_territry AS VARCHAR (20))

            if (iErr != 0)
                return;

            int i1 = 0;
            int iCnt = 0;

            SAPbobsCOM.Recordset oRS = null;

            try
            {
                oRS = (SAPbobsCOM.Recordset)oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (sSQL.Length == 0)
                {
                    // just remove the valid values
                    iCnt = oComboBox.ValidValues.Count;
                    if (iCnt < 1)
                       return;

                    for (i1 = iCnt - 1; i1 >= 0; i1 += -1)
                    {
                        oComboBox.ValidValues.Remove(i1, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                   return;
                }

                oRS.DoQuery(sSQL);
                if (oRS.RecordCount < 1)
                   return;

                for (i1 = 1; i1 <= oRS.RecordCount; i1++)
                {
                    try
                    {
                        oComboBox.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), oRS.Fields.Item(1).Value.ToString());
                    }
                    catch (Exception ex)
                    {
                        sTmp = "Entry already exists: " + oRS.Fields.Item(0).Value.ToString()+ " - " + oRS.Fields.Item(1).Value.ToString()+ Environment.NewLine;
                        sTmp = sTmp + sSQL;
                        oApp.MessageBox(sTmp + Environment.NewLine + "(ComboBoxAddVals)", 1, "Ok", "", "");
                    }
                    oRS.MoveNext();
                }


            }
            catch (Exception ex)
            {
                iErr = 1;
                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                // exit try ends up here, but the objects must be instantiated otherwise there will be and error.
                if (oRS != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                GC.Collect();
            }

        }

        public SAPbouiCOM.Form oAddForm(ref SAPbouiCOM.Application oApp, string sID, string sTitle, SAPbouiCOM.BoFormBorderStyle Border, int iTop, int iLeft, int iHeight, int iWidth, bool bVisible, ref int iErr)
        {
            SAPbouiCOM.Form functionReturnValue = default(SAPbouiCOM.Form);

            // routine to create a form

            // Border  --->
            //fbs_Sizable        sizable borders.  
            //fbs_Fixed          fixed-size borders.  
            //fbs_FixedNoTitle   fixed-size borders and no title.  
            //fbs_Floating       fixed-size borders, no title, and always on top.  

            functionReturnValue = null;

            if (iErr != 0) return null;

            SAPbouiCOM.FormCreationParams oCreationParams = null;

            try
            {
                oCreationParams = (SAPbouiCOM.FormCreationParams)oApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                oCreationParams.BorderStyle = Border;
                oCreationParams.UniqueID = sID;
                oCreationParams.FormType = sID;

                functionReturnValue = oApp.Forms.AddEx(oCreationParams);

                // Set the form properties 
                functionReturnValue.Title = sTitle;
                functionReturnValue.Top = iTop;
                functionReturnValue.Left = iLeft;
                functionReturnValue.ClientHeight = iHeight;
                functionReturnValue.ClientWidth = iWidth;
                functionReturnValue.Visible = bVisible;

            }
            catch (Exception ex)
            {
                iErr = 1;
                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                // exit try ends up here, but the objects must be instantiated otherwise there will be and error.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreationParams);
                GC.Collect();
            }
            return functionReturnValue;

        }

        public SAPbouiCOM.Button oAddButton(string FormUID, string sID, string sTitle, string sBaseItem, ref int iErr)
        {
            SAPbouiCOM.Button oRtn = default(SAPbouiCOM.Button);

            // sub to create a button on the form to the right of another object using the properties of that object
            // like adding another button next to the "cance;" button.

            if (iErr != 0) return null;

            int iSpace = 0;

            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Item oItm = null;


            try
            {
                oForm = oApp.Forms.Item(FormUID);

                // Get the space between the OK and Cancel buttons - if they exist
                try
                {
                    iSpace = oForm.Items.Item("2").Left - (oForm.Items.Item("1").Left + oForm.Items.Item("1").Width);
                }
                catch
                {
                    iSpace = 6;
                }

                oItm = oAddItem(FormUID, sID, sTitle, "", SAPbouiCOM.BoFormItemTypes.it_BUTTON, oForm.Items.Item(sBaseItem).Top,
                                oForm.Items.Item(sBaseItem).Left + oForm.Items.Item(sBaseItem).Width + iSpace, oForm.Items.Item(sBaseItem).Height, 
                                oForm.Items.Item(sBaseItem).Width, oForm.Items.Item(sBaseItem).FromPane,
                                oForm.Items.Item(sBaseItem).ToPane, oForm.Items.Item(sBaseItem).Enabled, oForm.Items.Item(sBaseItem).Visible, ref iErr);

                oRtn = (SAPbouiCOM.Button)oItm.Specific;

            }
            catch (Exception ex)
            {
                iErr = 1;
                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                // exit try ends up here, but the objects must be instantiated otherwise there will be an error.
                if (oForm != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                if (oItm != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oItm);

                GC.Collect();
            }
            return oRtn;
        }

        public SAPbouiCOM.Item oAddItem(string FormUID, string sID, string sTitle, string sLinkTo, SAPbouiCOM.BoFormItemTypes ObjType, 
                                        int iTop, int iLeft, int iHeight, int iWidth, int iFromPane,
                                        int iToPane, bool bActive, bool bVisible, ref int iErr)
        {
            SAPbouiCOM.Item oRtn = default(SAPbouiCOM.Item);

            // routine to create a UI object - FORM NAME

            // if the height or width is -1, then use the system default

            if (iErr != 0) return null;

            SAPbouiCOM.Form oForm = null;

            try
            {
                oForm = oApp.Forms.Item(FormUID);
                oRtn = oAddItem(ref oForm, sID, sTitle, sLinkTo, ObjType, 
                                iTop, iLeft, iHeight, iWidth, iFromPane,
                                iToPane, bActive, bVisible,ref iErr);

            }
            catch (Exception ex)
            {
                oRtn = null;
                iErr = 1;
                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                // exit try ends up here, but the objects must be instantiated otherwise there will be an error.
                if (oForm != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                GC.Collect();
            }
            return oRtn;

        }

        public SAPbouiCOM.Item oAddItem(ref SAPbouiCOM.Form oForm, string sID, string sTitle, string sLinkTo, SAPbouiCOM.BoFormItemTypes ObjType, 
                                        int iTop, int iLeft, int iHeight, int iWidth, int iFromPane,
                                        int iToPane, bool bActive, bool bVisible, ref int iErr)
        {
            SAPbouiCOM.Item oRtn = default(SAPbouiCOM.Item);

            // routine to create a UI object - ACTUAL FORM ID

            // if the height or width is -1, then use the system default


            if (iErr != 0) return null;

            //Dim oItm As SAPbouiCOM.Item = Nothing
            SAPbouiCOM.Button oBtn = null;
            SAPbouiCOM.StaticText oSta = null;
            SAPbouiCOM.EditText oEdt = null;
            SAPbouiCOM.EditText oExt = null;
            SAPbouiCOM.Grid oGrd = null;
            SAPbouiCOM.Matrix oMat = null;


            try
            {
                oRtn = oForm.Items.Add(sID, ObjType);
                oRtn.Top = iTop;
                oRtn.Left = iLeft;
                if (iHeight < 0)
                {
                    oRtn.Height = oApp.GetFormItemDefaultHeight((SAPbouiCOM.BoFormSizeableItemTypes)ObjType);
                }
                else
                {
                    oRtn.Height = iHeight;
                }
                if (iWidth < 0)
                {
                    oRtn.Width = oApp.GetFormItemDefaultWidth((SAPbouiCOM.BoFormSizeableItemTypes)ObjType);
                }
                else
                {
                    oRtn.Width = iWidth;
                }
                oRtn.Visible = bVisible;
                oRtn.Enabled = bActive;
                oRtn.FromPane = iFromPane;
                oRtn.ToPane = iToPane;
                if (!string.IsNullOrEmpty(sLinkTo)) oRtn.LinkTo = sLinkTo;

                switch (ObjType)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_BUTTON:
                        oBtn = (SAPbouiCOM.Button)oRtn.Specific;
                        if (!string.IsNullOrEmpty(sTitle))
                            oBtn.Caption = sTitle;
                        // mainly for "OK" and "Cancle" buttons
                        break;

                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                        oEdt = (SAPbouiCOM.EditText)oRtn.Specific;

                        break;
                    case (SAPbouiCOM.BoFormItemTypes.it_EXTEDIT):
                        oExt = (SAPbouiCOM.EditText)oRtn.Specific;

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                        oGrd = (SAPbouiCOM.Grid)oRtn.Specific;

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                        oMat = (SAPbouiCOM.Matrix)oRtn.Specific;

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_STATIC:
                        oSta = (SAPbouiCOM.StaticText)oRtn.Specific;
                        oSta.Caption = sTitle;

                        break;
                }

            }
            //catch (Exception ex)
            //{
            //    // Item already exists
            //    oRtn = oForm.Items.Item(sID);
            //}
            catch (Exception ex)
            {
                oRtn = null;
                iErr = 1;
                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                // exit try ends up here, but the objects must be instantiated otherwise there will be an error.
                //If Not oForm Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm)
                if ((oBtn != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBtn);
                if ((oSta != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSta);
                if ((oEdt != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdt);
                if ((oExt != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oExt);
                if ((oGrd != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrd);
                if ((oMat != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
                GC.Collect();
            }
            return oRtn;

        }
        #endregion
        #region "Utils"
        public void MessageUI(string sTxt)
        {
            //string sLine = null;
            //string sSub = null;
            //string sMsg = sTxt;
            //System.Diagnostics.StackFrame oFrame = null;

            //try
            //{
            //    oFrame = new System.Diagnostics.StackFrame(1, true);

            //    sSub = oFrame.GetMethod.Name;
            //    sLine = oFrame.GetFileLineNumber();
            //    sMsg = sMsg + Environment.NewLine + "(" + sSub + "     Line:" + sLine + ")";

            //    oApp.MessageBox(sMsg, 1, "Ok", "", "");

            //    // to get the entire stack trace:
            //    // System.Environment.StackTrace.ToString

            //}
            //catch (Exception ex)
            //{
            //    LIBDI.DImsg.MessageERR(ref ex);
            //}
            //finally
            //{
            //    if ((oFrame != null)) System.Runtime.InteropServices.Marshal.ReleaseComObject(oFrame);
            //    GC.Collect();
            //}

        }

        public void SetField(ref SAPbouiCOM.Form oForm, string sField, int iRow, string sCol, string sVal, ref int iErr)
        {
            //
            // combo boxes do not issue a "Validate" event only a "lost focus".
            // a "lost focus" event ONLY occures in an "after event".
            //
            SAPbouiCOM.PictureBox oPic = null;
            SAPbouiCOM.ComboBox oCmb = null;
            SAPbouiCOM.CheckBox oCkb = null;
            SAPbouiCOM.EditText oEdt = null;
            SAPbouiCOM.Column oCol = null;
            SAPbouiCOM.Matrix oMat = null;
            SAPbouiCOM.Matrix oGrd = null;
            //Dim oCels As SAPbouiCOM.Cells = Nothing
            //Dim oCel As SAPbouiCOM.Cell = Nothing

            if (iErr != 0)
                return;

            try
            {
                switch (oForm.Items.Item(sField).Type)
                {

                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON:
                        oEdt = (SAPbouiCOM.EditText)oForm.Items.Item(sField).Specific;
                        oEdt.Value = sVal;

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_PICTURE:
                        oPic = (SAPbouiCOM.PictureBox)oForm.Items.Item(sField).Specific;
                        oPic.Picture = sVal;

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item(sField).Specific;
                        oCmb.Select(sVal, SAPbouiCOM.BoSearchKey.psk_ByValue);

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                        oCkb = (SAPbouiCOM.CheckBox)oForm.Items.Item(sField).Specific;
                        oCkb.Checked = Convert.ToBoolean(sVal);

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                        oMat = (SAPbouiCOM.Matrix)oForm.Items.Item(sField).Specific;
                        oCol = oMat.Columns.Item(sCol);
                        switch (oCol.Type)
                        {
                            case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                            case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                            case SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON:
                                oEdt = (SAPbouiCOM.EditText)oMat.Columns.Item(sCol).Cells.Item(iRow).Specific;
                                oEdt.Value = sVal;

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_PICTURE:
                                oPic = (SAPbouiCOM.PictureBox)oMat.Columns.Item(sCol).Cells.Item(iRow).Specific;
                                oPic.Picture = sVal;

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                                oCmb = (SAPbouiCOM.ComboBox)oMat.Columns.Item(sCol).Cells.Item(iRow).Specific;
                                oCmb.Select(sVal, SAPbouiCOM.BoSearchKey.psk_ByValue);

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                                oCkb = (SAPbouiCOM.CheckBox)oMat.Columns.Item(sCol).Cells.Item(iRow).Specific;
                                oCkb.Checked = Convert.ToBoolean(sVal);
                                break;
                        }
                        break;
                }

            }
                // RHH - fix this ---------------------------------------------------------
            //catch (Exception ex)
            //{
            //    if (string.IsNullOrEmpty(sVal))
            //       return;
            //    // there is no valid value at this point
            //    iErr = 1;
            //    LIBDI.DImsg.MessageERR(ref ex);

            //}
            catch (Exception ex)
            {
                iErr = 1;
                //MessageBox.Show(ex.Message + vbCrLf + "(PutFieldValue)")
                 LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                if ((oPic != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPic);
                if ((oCmb != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmb);
                if ((oCkb != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCkb);
                if ((oEdt != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdt);
                if ((oCol != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCol);
                if ((oMat != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
                if ((oGrd != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrd);
                GC.Collect();
            }
        }
        public string sGetField(ref SAPbouiCOM.Form oForm, string sField, int iRow, string sCol, ref int iErr)
        {
            return sGetField(ref oForm, sField, iRow, sCol, "D", ref iErr);
        }
        public string sGetField(ref SAPbouiCOM.Form oForm, string sField, int iRow, string sCol, string sCmbBoxValOrDsc, ref int iErr)
        {
            string sRtn = null;
            //
            // combo boxes do not issue a "Validate" event only a "lost focus".
            // a "lost focus" event ONLY occures in an "after event".
            //

            // sCmbBoxValOrDsc   - V or D - if the field/ column is a combo box, this indicates to return the Value or the Description.

            SAPbouiCOM.PictureBox oPic = null;
            SAPbouiCOM.ComboBox oCmb = null;
            SAPbouiCOM.CheckBox oCkb = null;
            SAPbouiCOM.EditText oEdt = null;
            SAPbouiCOM.Column oCol = null;
            SAPbouiCOM.Matrix oMat = null;
            SAPbouiCOM.Grid oGrd = null;
            //Dim oCels As SAPbouiCOM.Cells = Nothing
            //Dim oCel As SAPbouiCOM.Cell = Nothing

            if (iErr != 0) return "";

            try
            {
                switch (oForm.Items.Item(sField).Type)
                {

                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON:
                        oEdt = (SAPbouiCOM.EditText)oForm.Items.Item(sField).Specific;
                        sRtn = oEdt.String;

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_PICTURE:
                        oPic = (SAPbouiCOM.PictureBox)oForm.Items.Item(sField).Specific;
                        sRtn = oPic.Picture;

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item(sField).Specific;
                        if (!string.IsNullOrEmpty(oCmb.Value))
                        {
                            sRtn = oCmb.Selected.Description;
                            if (sCmbBoxValOrDsc == "V")
                                sRtn = oCmb.Selected.Value;
                        }

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                        oCkb = (SAPbouiCOM.CheckBox)oForm.Items.Item(sField).Specific;
                        if (oCkb.Checked)
                        {
                            sRtn = "Y";
                        }
                        else
                        {
                            sRtn = "N";
                        }

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                        oMat = (SAPbouiCOM.Matrix)oForm.Items.Item(sField).Specific;
                        oCol = oMat.Columns.Item(sCol);
                        switch (oCol.Type)
                        {
                            case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                            case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                            case SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON:
                                //oEdt = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oEdt = (SAPbouiCOM.EditText)oMat.GetCellSpecific(sCol, iRow);
                                sRtn = oEdt.Value;

                                if (sCol == "37")
                                {
                                    //sTmp = oEdt.Value.GetTypeCode;
                                    decimal dtmp = Convert.ToDecimal(oEdt.Value);
                                    sTmp = "";
                                }

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_PICTURE:
                                //oPic = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oPic = (SAPbouiCOM.PictureBox)oMat.GetCellSpecific(sCol, iRow);
                                sRtn = oPic.Picture;

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                                //oCmb = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oCmb = (SAPbouiCOM.ComboBox)oMat.GetCellSpecific(sCol, iRow);
                                sRtn = oCmb.Selected.Description;

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                                //oCkb = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oCkb = (SAPbouiCOM.CheckBox)oMat.GetCellSpecific(sCol, iRow);
                                if (oCkb.Checked)
                                {
                                    sRtn = "Y";
                                }
                                else
                                {
                                    sRtn = "N";
                                }
                                break;
                        }

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                        oGrd = (SAPbouiCOM.Grid)oForm.Items.Item(sField).Specific;
                        sRtn = oGrd.DataTable.GetValue(sCol, iRow).ToString();
                        break;
                }

            }
            catch (Exception ex)
            {
                iErr = 1;
                                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                if ((oPic != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPic);
                if ((oCmb != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmb);
                if ((oCkb != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCkb);
                if ((oEdt != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdt);
                if ((oCol != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCol);
                if ((oMat != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
                if ((oGrd != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrd);
                GC.Collect();
                oPic = null;
                oCmb = null;
                oCkb = null;
                oEdt = null;
                oCol = null;
                oMat = null;
                oGrd = null;
            }
            return sRtn;

        }
        public decimal dGetField(ref SAPbouiCOM.Form oForm, string sField, int iRow, string sCol, ref int iErr)
        {
            decimal dRtn = default(decimal);
            //
            // combo boxes do not issue a "Validate" event only a "lost focus".
            // a "lost focus" event ONLY occures in an "after event".
            //
            SAPbouiCOM.ComboBox oCmb = null;
            SAPbouiCOM.CheckBox oCkb = null;
            SAPbouiCOM.EditText oEdt = null;
            SAPbouiCOM.Column oCol = null;
            SAPbouiCOM.Matrix oMat = null;
            SAPbouiCOM.Grid oGrd = null;

            dRtn = 0m;
            if (iErr != 0) return dRtn;

            try
            {
                switch (oForm.Items.Item(sField).Type)
                {

                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                        oEdt = (SAPbouiCOM.EditText)oForm.Items.Item(sField).Specific;
                        sTmp = oEdt.Value;
                        if (!string.IsNullOrEmpty(sTmp))
                            dRtn = Convert.ToDecimal(oEdt.Value);

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item(sField).Specific;
                        sTmp = oCmb.Selected.Description;
                        if (!string.IsNullOrEmpty(sTmp))
                            dRtn = Convert.ToDecimal(oCmb.Selected.Description);

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                        oCkb = (SAPbouiCOM.CheckBox)oForm.Items.Item(sField).Specific;
                        if (oCkb.Checked)
                        {
                            dRtn = 1m;
                        }
                        else
                        {
                            dRtn = 0m;
                        }

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                        oMat = (SAPbouiCOM.Matrix)oForm.Items.Item(sField).Specific;
                        oCol = oMat.Columns.Item(sCol);
                        switch (oCol.Type)
                        {
                            case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                            case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                                //oEdt = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oEdt = (SAPbouiCOM.EditText)oMat.GetCellSpecific(sCol, iRow);
                                sTmp = oEdt.Value;
                                if (!string.IsNullOrEmpty(sTmp))
                                    dRtn = Convert.ToDecimal(oEdt.Value);

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                                //oCmb = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oCmb = (SAPbouiCOM.ComboBox)oMat.GetCellSpecific(sCol, iRow);
                                sTmp = oCmb.Selected.Description;
                                if (!string.IsNullOrEmpty(sTmp))
                                    dRtn = Convert.ToDecimal(oCmb.Selected.Description);

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                                //oCkb = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oCkb = (SAPbouiCOM.CheckBox)oMat.GetCellSpecific(sCol, iRow);
                                if (oCkb.Checked)
                                {
                                    dRtn = 1m;
                                }
                                else
                                {
                                    dRtn = 0m;
                                }
                                break;
                        }

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                        oGrd = (SAPbouiCOM.Grid)oForm.Items.Item(sField).Specific;
                        sTmp = oGrd.DataTable.GetValue(sCol, iRow).ToString();
                        if (!string.IsNullOrEmpty(sTmp))
                            dRtn = Convert.ToDecimal(oGrd.DataTable.GetValue(sCol, iRow));
                        break;
                }
            }
            catch (Exception ex)
            {
                iErr = 1;
                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                if ((oCmb != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmb);
                if ((oCkb != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCkb);
                if ((oEdt != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdt);
                if ((oCol != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCol);
                if ((oMat != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
                if ((oGrd != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrd);
                GC.Collect();
            }
            return dRtn;

        }
        public int iGetField(ref SAPbouiCOM.Form oForm, string sField, int iRow, string sCol, ref int iErr)
        {
            int iRtn = 0;
            //
            // combo boxes do not issue a "Validate" event only a "lost focus".
            // a "lost focus" event ONLY occures in an "after event".
            //
            SAPbouiCOM.ComboBox oCmb = null;
            SAPbouiCOM.CheckBox oCkb = null;
            SAPbouiCOM.EditText oEdt = null;
            SAPbouiCOM.Column oCol = null;
            SAPbouiCOM.Matrix oMat = null;
            SAPbouiCOM.Grid oGrd = null;

            iRtn = 0;
            if (iErr != 0) return iRtn;

            try
            {
                switch (oForm.Items.Item(sField).Type)
                {

                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                        oEdt = (SAPbouiCOM.EditText)oForm.Items.Item(sField).Specific;
                        sTmp = oEdt.Value;
                        if (!string.IsNullOrEmpty(sTmp))
                            iRtn = Convert.ToInt32(oEdt.Value);

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        oCmb = (SAPbouiCOM.ComboBox)oForm.Items.Item(sField).Specific;
                        sTmp = oCmb.Selected.Description;
                        if (!string.IsNullOrEmpty(sTmp))
                            iRtn = Convert.ToInt32(oCmb.Selected.Description);

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                        oCkb = (SAPbouiCOM.CheckBox)oForm.Items.Item(sField).Specific;
                        if (oCkb.Checked)
                        {
                            iRtn = 1;
                        }
                        else
                        {
                            iRtn = 0;
                        }

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                        oMat = (SAPbouiCOM.Matrix)oForm.Items.Item(sField).Specific;
                        oCol = oMat.Columns.Item(sCol);
                        switch (oCol.Type)
                        {
                            case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                            case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                                //oEdt = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oEdt = (SAPbouiCOM.EditText)oMat.GetCellSpecific(sCol, iRow);
                                sTmp = oEdt.Value;
                                if (!string.IsNullOrEmpty(sTmp))
                                    iRtn = Convert.ToInt32(oEdt.Value);

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                                //oCmb = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oCmb = (SAPbouiCOM.ComboBox)oMat.GetCellSpecific(sCol, iRow);
                                sTmp = oCmb.Selected.Description;
                                if (!string.IsNullOrEmpty(sTmp))
                                    iRtn = Convert.ToInt32(oCmb.Selected.Description);

                                break;
                            case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                                //oCkb = oMat.Columns.Item(sCol).Cells.Item(iRow).Specific
                                oCkb = (SAPbouiCOM.CheckBox)oMat.GetCellSpecific(sCol, iRow);
                                if (oCkb.Checked)
                                {
                                    iRtn = 1;
                                }
                                else
                                {
                                    iRtn = 0;
                                }
                                break;
                        }

                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                        oGrd = (SAPbouiCOM.Grid)oForm.Items.Item(sField).Specific;
                        sTmp = oGrd.DataTable.GetValue(sCol, iRow).ToString();
                        if (!string.IsNullOrEmpty(sTmp))
                            iRtn = Convert.ToInt32(oGrd.DataTable.GetValue(sCol, iRow));
                        break;
                }
            }
            catch (Exception ex)
            {
                iErr = 1;
                                LIBDI.DImsg.MessageERR(ref ex);
            }
            finally
            {
                if ((oCmb != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmb);
                if ((oCkb != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCkb);
                if ((oEdt != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdt);
                if ((oCol != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCol);
                if ((oMat != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
                if ((oGrd != null))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrd);
                GC.Collect();
            }
            return iRtn;

        }
        //public string sGetFormUDF(ref string sSapFormID, ref short iAdj)
        //{

        //    // routine to take the current form handle and adjust the handle for a different form
        //    //
        //    //new property called "UDFFromUID" in the form object allowing you to directly get the UDFs form without any calculation.
        //    //
        //    // the reason for this routine is:
        //    //
        //    // on a standard B1 form, the handle will be something like "F_23".
        //    // the user defined fields(UDF) are actually on a seperate form and will have a handle of "F_24".
        //    // if the source of the event is from an UDF,  and the source or tagret field is on the standard form,
        //    // then the handle needs to be adjusted down by one to access the field(s) on the standard form.
        //    // same adjustment, except up by 1, then the event is from the standard form and either of the other fields
        //    // is on the UDF form.
        //    //

        //    int iNum = 0;
        //    iNum = Convert.ToInt32(String.Right(sSapFormID, Strings.Len(sSapFormID) - 2));
        //    iNum = Convert.ToInt32(sSapFormID.Substring(2));
        //    iNum = iNum + iAdj;
        //    return "F_" + iNum.ToString;

        //}
        //public SAPbouiCOM.Form oGetFormUDF(ref SAPbouiCOM.Form oActiveF, string sSub, ref int iErr)
        //{
        //    SAPbouiCOM.Form functionReturnValue = default(SAPbouiCOM.Form);

        //    // routine to take the current form handle and adjust the handle for a different form
        //    //
        //    //new property called "UDFFromUID" in the form object allowing you to directly get the UDFs form without any calculation.
        //    //
        //    // the reason for this routine is
        //    //
        //    // on a standard B1 form, the handle will be something like "F_23".
        //    // the user defined fields(UDF) are actually on a seperate form and will have a handle of "F_24".
        //    // if the source of the event is from an UDF,  and the source or tagret field is on the standard form,
        //    // then the handle needs to be adjusted down by one to access the field(s) on the standard form.
        //    // same adjustment, except up by 1, then the event is from the standard form and either of the other fields
        //    // is on the UDF form.
        //    //

        //    int iNum = 0;
        //    int iAdj = 0;
        //    string sFormID = null;

        //    functionReturnValue = null;
        //    if (iErr != 0)
        //        return;

        //    try
        //    {
        //        functionReturnValue = oActiveF;

        //        sFormID = oActiveF.UniqueID;

        //        if (Convert.ToInt32(oActiveF.TypeEx) < 0)
        //            iAdj = -1;
        //        if (Convert.ToInt32(oActiveF.TypeEx) > 0)
        //            iAdj = 1;
        //        if (iAdj == 0)
        //           return;

        //        iNum = Convert.ToInt32(sFormID.Substring(2));
        //        iNum = iNum + iAdj;
        //        sFormID = "F_" + iNum.ToString;

        //        functionReturnValue = oApp.Forms.Item(sFormID);

        //        // see it the form IDs are the same (139 and -139)
        //        if (Math.Abs(Convert.ToInt32(oActiveF.TypeEx)) != Math.Abs(Convert.ToInt32(functionReturnValue.TypeEx)))
        //        {
        //            //bubbleevent = False
        //            functionReturnValue = null;
        //            oApp.StatusBar.SetText("Form Numbers are not the same." + Environment.NewLine + "(" + sSub + " --> oGetFormUDF)", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //           return;
        //        }

        //        if (iAdj == -1 & functionReturnValue.Title != "Fields - Setup...")
        //        {
        //            oApp.StatusBar.SetText("UDF form's title is not 'Fields - Setup...'." + Environment.NewLine + "(" + sSub + " --> oGetFormUDF)", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //           return;
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        // -3012 occurs when the UDF form is not open
        //        if (Err.Number != -3012)
        //        {
        //            iErr = 1;
        //            //MessageBox.Show(ex.Message + vbCrLf + "(oGetFormUDF)")
        //                            LIBDI.DImsg.MessageERR(ref ex);
        //        }
        //    }

        //    GC.Collect();
        //    return functionReturnValue;

        //}
        #endregion
    }
}
