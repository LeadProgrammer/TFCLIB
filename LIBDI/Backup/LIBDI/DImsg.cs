using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;

namespace LIBDI
{
    public class DImsg
    {
        string sTmp = "";
        
        public static string MessageERR(int iErr, ref Exception oEX)
        {
            // entry point with an exception

            System.Diagnostics.StackTrace oTrace1 = new System.Diagnostics.StackTrace(oEX, true);
            // get the method and line number of the error
            System.Diagnostics.StackTrace oTrace2 = new System.Diagnostics.StackTrace(true);
            // get the method and line number of the calling methods
            System.Diagnostics.StackFrame oFrame = null;
            string sMethod = null;
            string sLine = null;
            string sMsg = "";

            int i = 0;

            try
            {
                sMsg = sMsg + "ERROR: " + iErr.ToString() + "    " + oEX.Message + Environment.NewLine;

                // get the error method and line number
                oFrame = oTrace1.GetFrame(oTrace1.FrameCount - 1);
                sLine = oFrame.GetFileLineNumber().ToString();
                sMethod = oFrame.GetMethod().Name + "                                                                 ";
                sMethod = sMethod.Substring(0, 40);

                sMsg = sMsg + "(" + sMethod + "     Line:" + sLine + ")";

                // now get the calling hierarchy
                for (i = 2; i <= oTrace2.FrameCount - 1; i++)
                {
                    oFrame = oTrace2.GetFrame(i);
                    sMethod = oTrace2.GetFrame(i).GetMethod().Name + "                                                  ";
                    sMethod = sMethod.Substring(0, 40);
                    sLine = oFrame.GetFileLineNumber().ToString();
                    // skip system routines
                    if (oFrame.GetFileLineNumber() > 0)
                    {
                        sMsg = sMsg + Environment.NewLine + "(" + sMethod + "Line:" + sLine + ")";
                    }
                }

                MessagePUT(sMsg);
                // display the message

            }
            catch (Exception ex)
            {

                //MsgBox("(MessageERR - oEX)     " + Err().Number.ToString() + "    " + Err().Description + Environment.NewLine + sMsg);
                MessageBox.Show("(MessageERR - oEX)     " + ex.Message + Environment.NewLine + sMsg);
            }

            //If Not oFrame Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oFrame)      ' causes an error
            //If Not oTrace1 Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oTrace1)    ' causes an error
            //If Not oTrace2 Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oTrace2)    ' causes an error
            GC.Collect();
            return sMsg;
        }

        public static string MessageERR(ref Exception oEX)
        {
            // entry point with an exception

            System.Diagnostics.StackTrace oTrace1 = new System.Diagnostics.StackTrace(oEX, true);
            // get the method and line number of the error
            System.Diagnostics.StackTrace oTrace2 = new System.Diagnostics.StackTrace(true);
            // get the method and line number of the calling methods
            System.Diagnostics.StackFrame oFrame = null;
            string sMethod = null;
            string sLine = null;
            string sMsg = "";

            int i = 0;

            try
            {
                sMsg = sMsg + "ERROR:  " + oEX.Message + Environment.NewLine;

                // get the error method and line number
                oFrame = oTrace1.GetFrame(oTrace1.FrameCount - 1);
                sLine = oFrame.GetFileLineNumber().ToString();
                sMethod = oFrame.GetMethod().Name + "                                                                 ";
                sMethod = sMethod.Substring(0, 40);

                sMsg = sMsg + "(" + sMethod + "     Line:" + sLine + ")";

                // now get the calling hierarchy
                for (i = 2; i <= oTrace2.FrameCount - 1; i++)
                {
                    oFrame = oTrace2.GetFrame(i);
                    sMethod = oTrace2.GetFrame(i).GetMethod().Name + "                                                  ";
                    sMethod = sMethod.Substring(0, 40);
                    sLine = oFrame.GetFileLineNumber().ToString();
                    // skip system routines
                    if (oFrame.GetFileLineNumber() > 0)
                    {
                        sMsg = sMsg + Environment.NewLine + "(" + sMethod + "Line:" + sLine + ")";
                    }
                }

                MessagePUT(sMsg);
                // display the message

            }
            catch (Exception ex)
            {

                //MsgBox("(MessageERR - oEX)     " + Err().Number.ToString() + "    " + Err().Description + Environment.NewLine + sMsg);
                MessageBox.Show("(MessageERR - oEX)     " + ex.Message + Environment.NewLine + sMsg);
            }

            //If Not oFrame Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oFrame)      ' causes an error
            //If Not oTrace1 Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oTrace1)    ' causes an error
            //If Not oTrace2 Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oTrace2)    ' causes an error
            GC.Collect();
            return sMsg;
        }

        public static string MessageERR(int iErr, string sTxt)
        {
            // routine to create and format the error message
            // entry point with out an exception

            string sMethod = null;
            string sLine = null;
            string sSub = null;
            string sMsg = "";
            System.Diagnostics.StackFrame oFrame = null;
            System.Diagnostics.StackTrace oTrace2 = new System.Diagnostics.StackTrace(true);
            // get the method and line number of the calling methods

            int i = 0;

            try
            {
                sMsg = sMsg + "ERROR: " + iErr.ToString() + "       " + sTxt;
                oFrame = new System.Diagnostics.StackFrame(1, true);
                sSub = oFrame.GetMethod().Name;
                sLine = oFrame.GetFileLineNumber().ToString();
                sMsg = sMsg + Environment.NewLine + "(" + sSub + "     Line:" + sLine + ")";

                // now get the calling hierarchy
                for (i = 2; i <= oTrace2.FrameCount - 1; i++)
                {
                    oFrame = oTrace2.GetFrame(i);
                    sMethod = oTrace2.GetFrame(i).GetMethod().Name + "                                                  ";
                    sMethod = sMethod.Substring(0, 40);
                    sLine = oFrame.GetFileLineNumber().ToString();
                    // skip system routines
                    if (oFrame.GetFileLineNumber() > 0)
                    {
                        sMsg = sMsg + Environment.NewLine + "(" + sMethod + "Line:" + sLine + ")";
                    }
                }
                MessagePUT(sMsg);
                // display the message

                // to get the entire stack trace:
                // System.Environment.StackTrace.ToString

            }
            catch (Exception ex)
            {
                //Interaction.MsgBox("(MessageERR - sTxt)     " + Err().Number.ToString() + "    " + Err().Description + Environment.NewLine + sMsg);
                MessageBox.Show("(MessageERR - sTxt)     " + ex.Message + Environment.NewLine + sMsg);
            }
            finally
            {
                //If Not oFrame Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oFrame)
                GC.Collect();
            }
            return sMsg;
        }

        public static string MessageERR(string sTxt)
        {
            // routine to create and format the error message
            // entry point with out an exception

            // Then Throw New Exception("Matrix is empty.")

            int i = 0;

            string sMethod = null;
            string sLine = null;
            string sSub = null;
            string sMsg = "";
            System.Diagnostics.StackFrame oFrame = null;
            System.Diagnostics.StackTrace oTrace2 = new System.Diagnostics.StackTrace(true);
            // get the method and line number of the calling methods


            try
            {
                sMsg = sTxt;
                oFrame = new System.Diagnostics.StackFrame(1, true);
                sSub = oFrame.GetMethod().Name;
                sLine = oFrame.GetFileLineNumber().ToString();
                sMsg = sMsg + Environment.NewLine + "(" + sSub + "     Line:" + sLine + ")";

                // now get the calling hierarchy
                for (i = 2; i <= oTrace2.FrameCount - 1; i++)
                {
                    oFrame = oTrace2.GetFrame(i);
                    sMethod = oTrace2.GetFrame(i).GetMethod().Name + "                                                  ";
                    sMethod = sMethod.Substring(0, 40);
                    sLine = oFrame.GetFileLineNumber().ToString();
                    // skip system routines
                    if (oFrame.GetFileLineNumber() > 0)
                    {
                        sMsg = sMsg + Environment.NewLine + "(" + sMethod + "Line:" + sLine + ")";
                    }
                }

                MessagePUT(sMsg);
                // display the message

                // to get the entire stack trace:
                // System.Environment.StackTrace.ToString

            }
            catch (Exception ex)
            {
                //Interaction.MsgBox("(MessageERR - sTxt)     " + Err().Number.ToString() + "    " + Err().Description + Environment.NewLine + sMsg);
                MessageBox.Show("(MessageERR - sTxt)     " + ex.Message + Environment.NewLine + sMsg);
            }
            finally
            {
                //If Not oFrame Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oFrame)
                GC.Collect();
            }
            return sMsg;
        }

        public static Boolean MessagePUT(string sTxt)
        {
            // display and\or log the Message

            string sMsg = "";
            SAPbobsCOM.Recordset oRS1 = null;

            try
            {
                // see what to do with the error message

                // OCOMP MAY NOT EXIST IF BEING CALLED FROM EXTERNAL PROGRAM - ALSO MAY NOT KNOW THE DATABASE YET - NEED TO A SETMSG SUB TO SET THESE VARS

                sMsg = sTxt;

                //oRS1 = oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                //oRS1.DoQuery("SELECT U_Value FROM [" + scConfig + "] WITH (NOLOCK) WHERE Name = N'Message Level'")
                //sTmp = oRS1.Fields.Item(0).Value

                //If sTmp.ToUpper = "NONE" Then Exit Sub
                //If sTmp.ToUpper = "" Then sMsg = sTxt
                //If sTmp.ToUpper = "MSG" Then sMsg = sTxt
                //If sTmp.ToUpper = "TRACE" Then sMsg = sTxt + Environment.NewLine + System.Environment.StackTrace.ToString


                //if (scUIDI == "UI")
                MessageSHOW(sMsg);
                // show the message

                //MessageLOG(sMsg);
                // write the message to the log (if specified)

            }
            catch (Exception ex)
            {
                //if (TFCLIB_DI.scUIDI == "UI")
                //{
                //Interaction.MsgBox("(MessagePUT)     " + ex.Number.ToString() + "    " + Err().Description + Environment.NewLine + sTxt);
                MessageBox.Show("(MessagePUT)     " + ex.Message + Environment.NewLine + sMsg);
                //}
                //MessageLOG("(MessagePUT)     " + Err().Number.ToString() + "    " + Err().Description + Environment.NewLine + sTxt);
            }

            if ((oRS1 != null)) System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS1);
            GC.Collect();
            return true;
        }

        public static Boolean MessageSHOW(string sTxt)
        {
            // display the Message

            // if in DI mode, do not show the message
            if (LIBDI.scUIDI == "DI") return true;

            IntPtr ptr = MicroSoftWindows.GetForegroundWindow();
            MicroSoftWindows.WindowWrapper oWindow = new MicroSoftWindows.WindowWrapper(ptr);

            try
            {
                MicroSoftWindows.SetForegroundWindow(ptr);

                MessageBox.Show(oWindow, sTxt, "ERROR IN ADD-ON");

            }
            catch (Exception ex)
            {
                //if (TFCLIB_DI.scUIDI == "UI")
                //{
                //Interaction.MsgBox("(MessageSHOW)     " + Err().Number.ToString() + "    " + Err().Description + Environment.NewLine + sTxt);
                MessageBox.Show("(MessageSHOW)     " + ex.Message + Environment.NewLine + sTxt);
                //}
                //else
                //{
                //    //MessageLOG("(MessageSHOW)     " + Err().Number.ToString() + "    " + Err().Description + Environment.NewLine + sTxt);
                //    MessageBox.Show("(MessageSHOW)     " + ex.Message + Environment.NewLine + sTxt);
                //}
            }

            //If Not oWindow Is Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oWindow)    ' causes an error
            GC.Collect();
            return true;
        }

        public static Boolean MessageEMAIL(string sTxt)
        {
            try
            {
                // get the default Recipients


                return true;
            }
            catch (SmtpFailedRecipientException ex)
            {
                return false;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static Boolean MessageEMAIL(string sRecipients, string sTxt)
        {

            // E Mail Message

            if (string.IsNullOrEmpty(sRecipients)) return false;

            string SMTPHost = "EXCHANGE.fruitco.local";

            System.Net.NetworkCredential cred = new System.Net.NetworkCredential();
            cred.Domain = "FRUITCO";
            cred.UserName = "ITSupport";
            cred.Password = "qp?n-ytu?";

            SmtpClient client = new SmtpClient(SMTPHost);
            client.UseDefaultCredentials = false;
            client.Credentials = cred;

            string fromAddress = "email@TheFruitCompany.com";
            string bcc = "";
            string toAddress = sRecipients;
            string subject = "Error Notification";

            System.Net.Mail.MailMessage eMsg = new System.Net.Mail.MailMessage();

            try
            {
                eMsg.IsBodyHtml = true;
                eMsg.From = new System.Net.Mail.MailAddress(fromAddress);

                eMsg.To.Clear();
                eMsg.To.Add(toAddress);
                eMsg.Bcc.Clear();
                eMsg.Bcc.Add(bcc);
                eMsg.Subject = subject;
                eMsg.Body = sTxt;

                client.Send(eMsg);

                return true;
            }
            catch (SmtpFailedRecipientException ex)
            {
                return false;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
