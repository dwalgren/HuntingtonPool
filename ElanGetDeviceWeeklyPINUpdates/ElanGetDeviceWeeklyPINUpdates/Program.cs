using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Net.Mail;

namespace ElanGetDeviceWeeklyPINUpdates
{
    class Program
    {
        static void Main(string[] args)
        {
            ProcessDevices();
        }
        static void ProcessDevices()
        {
            DateTime dtNow = DateTime.Now;
            StringBuilder sb = new StringBuilder();
            string connectionString = ConfigurationManager.ConnectionStrings["elanConnectDevices"].ConnectionString;
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                sb.Append("<p><strong>Viking PIN Updates</strong><br />&nbsp;</p>");
                sb.Append("<p>Date: " + dtNow.ToString() + "</p>");

                string query = "exec [dbo].[GetDevicesPINUpdatesToNotify]";
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    DataTable dtDevices = new DataTable("Devices");
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    String LastProcessedID = string.Empty;
                    String LastProcessedMACAddress = string.Empty;
                    conn.Open();
                    da.Fill(dtDevices);
                    conn.Close();

                    if (dtDevices.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtDevices.Rows)
                        {
                            if (dr["MACAddress"] != null)
                            {
                                sb.Append(dr["MACAddress"].ToString() + ", " + dr["PIN"].ToString() + "<br />");
                                LastProcessedID = dr["ID"].ToString();
                                LastProcessedMACAddress = dr["MACAddress"].ToString();
                            }
                        }
                        sb.Append("<p><strong>Total: " + dtDevices.Rows.Count.ToString() + "</strong></p>");

                        string query2 = "exec [dbo].[UpdateDeviceUpdatesNotification] " + LastProcessedID + ", '" + LastProcessedMACAddress + "'";
                        using (SqlCommand cmd2 = new SqlCommand(query2, conn))
                        {
                            conn.Open();
                            int result = cmd2.ExecuteNonQuery();
                            conn.Close();
                        }
                    }
                    else
                    {
                        sb.Append("<p>No devices with PIN updates since last run</p>");
                    }
                }

                SendEmail(sb);
            }
        }
        private static void SendEmail(StringBuilder sb)
        {
            MailMessage mail = new MailMessage();
            //SmtpClient smtpServer = new SmtpClient("smtpout.secureserver.net");
            SmtpClient smtpServer = new SmtpClient("smtp.live.com");
            smtpServer.UseDefaultCredentials = false;
            smtpServer.Credentials = new System.Net.NetworkCredential("dwalgren@msn.com", "GoF0rward!");
            smtpServer.Port = 587; //25
            smtpServer.EnableSsl = true;
            mail.IsBodyHtml = true;
            mail.From = new MailAddress("dwalgren@msn.com");

            // Get the Recipients from App.config
            var recipients = new List<string>(ConfigurationManager.AppSettings["emailRecipients"].Split(new char[] { ';' }));
            foreach (string recipient in recipients)
            {
                if (!string.IsNullOrEmpty(recipient))
                {
                    mail.To.Add(recipient);
                }
            }
            //mail.To.Add("dwalgren@msn.com");
            //mail.To.Add("doug@boomerangcomputer.com");
            //mail.To.Add("jking@elanindustries.com");

            mail.Subject = "Elan Device Weekly PIN Update Notification";
            mail.Body = sb.ToString();
            smtpServer.Send(mail);
            return;
        }
    }
}
