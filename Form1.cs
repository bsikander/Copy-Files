using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Xml;
using System.Xml.Xsl;

namespace CopyFilesProj
{
    public partial class Form1 : Form
    {
        bool mNoCopy = false;

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// This function handles the click event of the Start button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "Copying Process Started!";

            string lFileName = ReadConfigFile("Filename");
            string lSourcePath = ReadConfigFile("Source");
            string lDestinationPath = ReadConfigFile("Target");
            DateTime lTimeStamp;

            if (lFileName.Trim() == string.Empty || lSourcePath.Trim() == string.Empty || lDestinationPath.Trim() == string.Empty)
            {
                //MessageBox.Show("No FileName or Source or Target listed in the App.Config file", "Error");
                lblStatus.Text = "No FileName/Source/Target listed in Config file";
                return;
            }

            //check the timestamp from the output.xml file
            lblStatus.Text = "Comparing File and Output.xml times!";
            if (CompareTimeStamps())
            {
                lblStatus.Text = "Copying File!";
                if (CopyFile(lFileName, lSourcePath, lDestinationPath, out lTimeStamp))
                {
                    //Write timestamp to output.xml
                    WriteTimeStampToXML(lTimeStamp);

                    lblStatus.Text = "Sending Success Email!";
                    //Email Success
                    SendEmail(ReadConfigFile("Emails"), ReadConfigFile("EmailSuccess"));
                }
                else
                {
                    //MessageBox.Show("Some Error occured during the copy process. File not copied.","Error");                    
                    lblStatus.Text = "Sending Fail Email!";
                    SendEmail(ReadConfigFile("Emails"), ReadConfigFile("EmailFail"));
                }
            }
            else if (mNoCopy == true)
            {
                // send no copy email
                lblStatus.Text = "Sending No Copy Email!";
                SendEmail(ReadConfigFile("Emails"), ReadConfigFile("EmailNoCopy"));
            }
        }

        /// <summary>
        /// This function eliminates all the extra things are in the DateTime extracted from the file.
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private DateTime RoundToSecond(DateTime dt)
        {
            return new DateTime(dt.Year, dt.Month, dt.Day,
                                dt.Hour, dt.Minute, dt.Second);
        }

        /// <summary>
        /// This function compares the time and returns a boolean.
        /// </summary>
        /// <returns></returns>
        private bool CompareTimeStamps()
        {
            string lSourceFilePath = System.IO.Path.Combine(ReadConfigFile("Source"), ReadConfigFile("Filename"));
            DateTime lXMLTimeStamp = ReadTimeStampFromXML();
            DateTime lFileTimeStamp;
            DateTime lNewDate;

            if (lXMLTimeStamp == DateTime.MinValue)
            {
                return true;
            }

            if (System.IO.Directory.Exists(ReadConfigFile("Source")))
            {
                lFileTimeStamp = System.IO.File.GetLastWriteTime(lSourceFilePath);

                //This function removes all the extra things added in the date like milliseconds/ticks
                lFileTimeStamp = RoundToSecond(lFileTimeStamp);

                int lRetVal = DateTime.Compare(lFileTimeStamp, lXMLTimeStamp);
                if (lFileTimeStamp > lXMLTimeStamp)
                    return true;
                else
                {
                    mNoCopy = true;
                    return false;
                }
            }
            else
            {
                MessageBox.Show("Source directory doesnot exist", "Error");
                return false;
            }
        }

        /// <summary>
        /// This function writes the TimeStamp to the XML file
        /// </summary>
        /// <param name="pTimeStamp"></param>
        private void WriteTimeStampToXML(DateTime pTimeStamp)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(ReadConfigFile("XMLFilePath"));
            XmlElement root = doc.DocumentElement;
            XmlNode lNode = root.SelectSingleNode("//timeStamp");
            lNode.InnerText = pTimeStamp.ToString();
            doc.Save(ReadConfigFile("XMLFilePath"));
        }

        /// <summary>
        /// This function reads the time stamp from the XML.
        /// </summary>
        /// <returns></returns>
        private DateTime ReadTimeStampFromXML()
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(ReadConfigFile("XMLFilePath"));
            XmlElement root = doc.DocumentElement;
            XmlNode lNode = root.SelectSingleNode("//timeStamp");

            if (lNode.InnerText == string.Empty)            
                return DateTime.MinValue;            
            else
                return Convert.ToDateTime(lNode.InnerText);
        }

        /// <summary>
        /// This function copies the file from source to destination and return the timestamp of the file
        /// </summary>
        /// <param name="pFileName">Filename </param>
        /// <param name="pSource">Source directory</param>
        /// <param name="pDestination">Target directory</param>
        /// <param name="pModifiedDate">modified date that is returned as the out</param>
        /// <returns></returns>
        private bool CopyFile(string pFileName, string pSource, string pDestination, out DateTime pModifiedDate)
        {
            string lSourceFilePath = System.IO.Path.Combine(pSource, pFileName);
            string lDestFilePath = System.IO.Path.Combine(pDestination, pFileName);

            if (System.IO.Directory.Exists(pDestination) && System.IO.Directory.Exists(pSource))
            {
                // To copy a file to another location and 
                // overwrite the destination file if it already exists.
                System.IO.File.Copy(lSourceFilePath, lDestFilePath, true);

                pModifiedDate = System.IO.File.GetLastWriteTime(lSourceFilePath);

                return System.IO.File.Exists(lDestFilePath);
            }
            else
            {
                MessageBox.Show("Source or Target Directory doesnot Exist", "Error");
                pModifiedDate = DateTime.MinValue;
                return false;
            }
        }

        /// <summary>
        /// This function reads the app.config. 
        /// </summary>
        /// <param name="pSettingName">Key name for the setting</param>
        /// <returns></returns>
        private string ReadConfigFile(string pSettingName)
        {
            return ConfigurationManager.AppSettings[pSettingName];
        }

        /// <summary>
        /// This functino send emails
        /// </summary>
        /// <param name="pToAddress"></param>
        /// <param name="pSubject"></param>
        private void SendEmail(string pToAddress, string pSubject)
        {
            try
            {   
                string lFromEmailId = ReadConfigFile("FromEmailId");
                string lUsername = ReadConfigFile("UserName");
                string lPassword = ReadConfigFile("Password");

                System.Net.Mail.MailMessage lMessage = new System.Net.Mail.MailMessage();

                string[] lToAddress = pToAddress.Split(';');
                foreach (string lAddr in lToAddress)
                {
                    lMessage.To.Add(lAddr);
                }
                                    
                lMessage.Subject = pSubject;
                lMessage.From = new System.Net.Mail.MailAddress(lFromEmailId);
                lMessage.Body = string.Empty;

                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient(ReadConfigFile("SMTPServer"), Convert.ToInt32(ReadConfigFile("SMTPPort")));
                smtp.EnableSsl = true;
                smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new System.Net.NetworkCredential(lUsername, lPassword);
                smtp.Send(lMessage);             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }        
    }
}
