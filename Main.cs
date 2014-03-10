using System;
using System.Collections.Generic;
using System.Text;
using PureCM.Server;
using PureCM.Client;
using System.Net;
using System.Net.Mail;
using System.Xml.Linq;
using System.IO;
using System.Security;
using System.Text.RegularExpressions;
using System.Web;
using System.Diagnostics;

namespace Plugin_EMailTasks
{
    [EventHandlerDescription("Plugin that allows for emailing task descriptions to users")]
    public class EMailTasksPlugin : PureCM.Server.Plugin
    {
        public override bool OnStart( XElement oConfig, Connection conn)
        {
            LogInfo("Starting Email Tasks Plugin");

            m_oFromAddr = new MailAddress( oConfig.Element("From").Value );
            m_listRecipients = new List<MailAddress>();

            try
            {
                foreach (XElement recipient in oConfig.Element("Recipients").Elements("Recipient"))
                {
                    m_listRecipients.Add(new MailAddress(recipient.Value));
                }
            }
            catch (Exception e)
            {
                LogWarning(String.Format("Handled exception when parsing recipient list - {0}", e.Message));
            }

            if ( m_listRecipients.Count == 0 )
            {
                LogWarning("No recipients configured for EMailChangesets plugin");
                return false;
            }

            m_strSMTPServer = oConfig.Element("SMTPServer").Value;

            XElement oSMTPUser = oConfig.Element("SMTPUser");

            if ( oSMTPUser != null )
            {
                m_oCredentials = new NetworkCredential(oSMTPUser.Value, oConfig.Element("SMTPPassword").Value);
            }
            else
            {
                m_oCredentials = null;
            }

            try
            {
                m_oReposPattern = new Regex(oConfig.Element("ReposPattern").Value);
            }
            catch (Exception e)
            {
                LogWarning(String.Format("Invalid ReposPattern value for EMailChangesets plugin: {0}", e.Message));
                m_oReposPattern = null;
            }

            m_strSubjectFormat = GetConfigString(oConfig, "SubjectFormat", "Task {0} has been created in repository {1}");

            conn.OnProjectItemCreated = OnProjectItemCreated;

            return true;
        }

        public override void OnStop()
        {
            LogInfo("Stopping Email Tasks Plugin");
        }

        private void SendEmail(string strRecipient, string strSubject, string strBody)
        {
            LogInfo(string.Format("Sending '{0}'", strSubject));

            var msg = new MailMessage();

            msg.From = m_oFromAddr;

            if (strRecipient.Length > 0)
            {
                LogInfo(string.Format("Sending to '{0}'", strRecipient));
                msg.To.Add(strRecipient);
            }
            else
            {
                foreach (MailAddress addr in m_listRecipients)
                {
                    LogInfo(string.Format("Sending to '{0}'", addr));
                    msg.To.Add(addr);
                }
            }

            msg.Subject = strSubject;
            msg.Body = strBody;
            msg.IsBodyHtml = true;

            var client = new SmtpClient(m_strSMTPServer);

            if ( m_oCredentials != null )
            {
                client.UseDefaultCredentials = false;
                client.Credentials = m_oCredentials;
            }

            client.Send(msg);
        }

        private void OnProjectItemCreated(ProjectItemEvent evt)
        {
            LogInfo(string.Format("Task {0} has been assigned", evt.ProjectItemID));

            if ((m_oReposPattern != null) &&
                (evt.Repository != null) &&
                !m_oReposPattern.IsMatch(evt.Repository.Name))
            {
                return;
            }

            var item = evt.Repository.ProjectItemById(evt.ProjectItemID);

            String name = evt.Repository.ProjectItemById(evt.ProjectItemID).Name;

            SendEmail("", String.Format(m_strSubjectFormat, name, evt.Repository.Name), "");
        }

        private bool GetConfigBool(XElement oConfig, String strName, bool bDefault)
        {
            XElement oElt = oConfig.Element(strName);
            bool bRet = bDefault;

            if ( (oElt != null) && ( oElt.Value != null ) )
            {
                bRet = bool.Parse(oElt.Value);
            }

            LogInfo(string.Format("Config '{0}'={1}", strName, bRet));

            return bRet;
        }

        private int GetConfigInt(XElement oConfig, String strName, int nDefault)
        {
            XElement oElt = oConfig.Element(strName);
            int nRet = nDefault;

            if ((oElt != null) && (oElt.Value != null))
            {
                nRet = int.Parse(oElt.Value);
            }
            return nRet;
        }

        private String GetConfigString(XElement oConfig, String strName, String strDefault)
        {
            XElement oElt = oConfig.Element(strName);
            String strRet = strDefault;

            if ((oElt != null) && (oElt.Value != null))
            {
                strRet = oElt.Value;
            }
            return strRet;
        }

        private MailAddress m_oFromAddr;
        private List<MailAddress> m_listRecipients;
        private String m_strSMTPServer;
        private NetworkCredential m_oCredentials;
        private String m_strSubjectFormat;
        private Regex m_oReposPattern;
    }
}
