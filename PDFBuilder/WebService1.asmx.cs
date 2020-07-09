using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Services;
using System.Web.Script.Serialization;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.IO;
using System.Drawing.Printing;
using System.Drawing;
using System.Diagnostics;
using PDFBuilder.rs2005;
using PDFBuilder.rsExec;


namespace PDFBuilder
{
    /// <summary>
    /// Summary description for WebService1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class WebService1 : System.Web.Services.WebService
    {
        public class openDB
        {
            public static SqlConnection ddb()
            {
                string connectionString = ConfigurationManager.ConnectionStrings["AGProd"].ConnectionString;
                SqlConnection conn = new SqlConnection(connectionString);
                return conn;
            }
        }

        [WebMethod]
        public string HelloWorld()
        {
            
            return "Hello World";
        }

        [WebMethod]
        public string GetLogonInfo()
        {
            string principal = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string authType = User.Identity.AuthenticationType;
            string usr = User.Identity.Name;
            string isAuth = User.Identity.IsAuthenticated.ToString();
            return "user " + usr + " is authenticated " + isAuth + " principal name " + principal + " with type " + authType;
        }

        [WebMethod]
        public string MakePDF(string link)
        {
            //  Get Required Parameters
            string ReportFolder = "Bogus Report";
            string ReportName = "RHS Report";
            string OrderNumber = "57654";
            string ToEmail = "";
            string Subject = "";

            rs2005.ReportingService2005 rs = new ReportingService2005();
            rsExec.ReportExecutionService rsExecService = new ReportExecutionService();
            rs.Credentials = System.Net.CredentialCache.DefaultCredentials;
            rsExecService.Credentials = System.Net.CredentialCache.DefaultCredentials;
            rs.Url = System.Configuration.ConfigurationManager.AppSettings["rs2005"];
            rsExecService.Url = System.Configuration.ConfigurationManager.AppSettings["rsExecService"];

            string historyID = null;
            string deviceInfo = null;
            string format = "PDF";
            byte[] results;
            string encoding = string.Empty;
            string mimType = string.Empty;
            string extension = string.Empty;
            rsExec.Warning[] warnings = null;
            string[] streamIDs = null;
            string OutputFileDirectory = System.Configuration.ConfigurationManager.AppSettings["Email_Output_File_Directory"];
            if (!Directory.Exists(OutputFileDirectory))
            {
                Directory.CreateDirectory(OutputFileDirectory);
            }

            string fileName = "email_" + ReportName;
            string filePath = @"" + OutputFileDirectory + fileName + ".pdf";
            string _reportName = @"/" + ReportFolder + "/" + ReportName;
            bool _forRendering = false;
            string _historyID = null;
            rs2005.ParameterValue[] _values = null;
            rs2005.DataSourceCredentials[] _credentials = null;
            rs2005.ReportParameter[] _parameters = null;
            
            _parameters = rs.GetReportParameters(_reportName, _historyID, _forRendering, _values, _credentials);
            rsExec.ExecutionInfo ei = rsExecService.LoadReport(_reportName, historyID);

            //  Get Parameter List
            string rptname = ReportName.Replace("+", " ");

            rsExec.ParameterValue[] parameters = new rsExec.ParameterValue[2];
            parameters[0] = new rsExec.ParameterValue();
            parameters[0].Label = "RHS Job ID";
            parameters[0].Name = "ID";
            parameters[0].Value = OrderNumber;

            Subject = "RHS Job Test" + OrderNumber;
            rsExecService.SetExecutionParameters(parameters, "en-us");
            results = rsExecService.Render(format, deviceInfo, out extension, out encoding, out mimType, out warnings, out streamIDs);

            //This is to get the datetime of save and removes special characters to create unique name
            string datetimecode = DateTime.Now.ToString();
            datetimecode = datetimecode.Replace("/", "");
            datetimecode = datetimecode.Replace(":", "");
            datetimecode = datetimecode.Replace(" ", "");

            filePath = @"" + OutputFileDirectory + fileName + "_" + OrderNumber + "_" + datetimecode + ".pdf";


            if (File.Exists(filePath))
            {
                Random rnd = new Random();
                int ending = rnd.Next(1000);
                filePath = @"" + OutputFileDirectory + fileName + "_" + OrderNumber + "_" + datetimecode + "_" + ending + ".pdf";
            }

            using (FileStream stream = File.OpenWrite(filePath))
            {
                stream.Write(results, 0, results.Length);
            }

            string Body = "";

            if (ToEmail == "NONE")
            {
                ToEmail = System.Configuration.ConfigurationManager.AppSettings["BadEmailTo"];
                Subject = "RHS Order " + OrderNumber;
                Body = "Test PDF " + OrderNumber + ". ";
            }

            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LIVEDB"].ToString());
            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "PP_SendEmailwithAttachment";
            SqlParameterCollection sqlParams = cmd.Parameters;
            sqlParams.AddWithValue("@userEmail", ToEmail);
            sqlParams.AddWithValue("@profileName", System.Configuration.ConfigurationManager.AppSettings["EmailSendingProfileName"]);
            sqlParams.AddWithValue("@subject", Subject);
            sqlParams.AddWithValue("@body", Body);
            sqlParams.AddWithValue("@attachment", filePath);
            sqlParams.AddWithValue("@BBCEmail", System.Configuration.ConfigurationManager.AppSettings["BBCEmails"]);
            sqlParams.AddWithValue("@fromEmails", System.Configuration.ConfigurationManager.AppSettings["EmailsFrom"]);
            sqlParams.AddWithValue("@orderID", System.Configuration.ConfigurationManager.AppSettings["EmailsFrom"]);
            sqlParams.AddWithValue("@Confirmation", System.Configuration.ConfigurationManager.AppSettings["EmailsFrom"]);
            conn.Open();
            cmd.Connection = conn;
            cmd.ExecuteReader();
            cmd.Dispose();
            conn.Close();
            conn.Dispose();

            return "done";
        }

        [WebMethod]
        public void BuildPDF()
        {
            try
            {
                HttpContext postedContext = HttpContext.Current;
                string ID = postedContext.Request.Form["ID"].ToString();
                string Email = postedContext.Request.Form["Email"].ToString();

                //  Get Required Parameters
                string ReportFolder = "Bogus Report";
                string ReportName = "BuybackletterPDF";
                string ToEmail = Email;
                string Subject = "";

                rs2005.ReportingService2005 rs = new ReportingService2005();
                rsExec.ReportExecutionService rsExecService = new ReportExecutionService();
                rs.Credentials = System.Net.CredentialCache.DefaultCredentials;
                rsExecService.Credentials = System.Net.CredentialCache.DefaultCredentials;
                rs.Url = System.Configuration.ConfigurationManager.AppSettings["rs2005"];
                rsExecService.Url = System.Configuration.ConfigurationManager.AppSettings["rsExecService"];

                string historyID = null;
                string deviceInfo = null;
                string format = "PDF";
                byte[] results;
                string encoding = string.Empty;
                string mimType = string.Empty;
                string extension = string.Empty;
                rsExec.Warning[] warnings = null;
                string[] streamIDs = null;
                string OutputFileDirectory = System.Configuration.ConfigurationManager.AppSettings["Email_Output_File_Directory"];
                if (!Directory.Exists(OutputFileDirectory))
                {
                    Directory.CreateDirectory(OutputFileDirectory);
                }

                //string fileName = "email_" + ReportName;
                string fileName = "HudsonBuybackLetter";
                string filePath = @"" + OutputFileDirectory + fileName + ".pdf";
                string _reportName = @"/" + ReportFolder + "/" + ReportName;
                bool _forRendering = false;
                string _historyID = null;
                rs2005.ParameterValue[] _values = null;
                rs2005.DataSourceCredentials[] _credentials = null;
                rs2005.ReportParameter[] _parameters = null;

                _parameters = rs.GetReportParameters(_reportName, _historyID, _forRendering, _values, _credentials);
                rsExec.ExecutionInfo ei = rsExecService.LoadReport(_reportName, historyID);

                //  Get Parameter List
                string rptname = ReportName.Replace("+", " ");

                rsExec.ParameterValue[] parameters = new rsExec.ParameterValue[2];
                parameters[0] = new rsExec.ParameterValue();
                parameters[0].Label = "RHS Job ID";
                parameters[0].Name = "ID";
                parameters[0].Value = ID;

                Subject = "Hudson Technologies Buyback Letter";
                rsExecService.SetExecutionParameters(parameters, "en-us");
                results = rsExecService.Render(format, deviceInfo, out extension, out encoding, out mimType, out warnings, out streamIDs);

                //This is to get the datetime of save and removes special characters to create unique name
                string datetimecode = DateTime.Now.ToString();
                datetimecode = datetimecode.Replace("/", "");
                datetimecode = datetimecode.Replace(":", "");
                datetimecode = datetimecode.Replace(" ", "");

                filePath = @"" + OutputFileDirectory + fileName + "_" + ID + "_" + datetimecode + ".pdf";
                //filePath = @"" + OutputFileDirectory + fileName + ".pdf";


                if (File.Exists(filePath))
                {
                    Random rnd = new Random();
                    int ending = rnd.Next(1000);
                    filePath = @"" + OutputFileDirectory + fileName + "_" + ID + "_" + datetimecode + "_" + ending + ".pdf";
                    //filePath = @"" + OutputFileDirectory + fileName + ".pdf";
                }

                using (FileStream stream = File.OpenWrite(filePath))
                {
                    stream.Write(results, 0, results.Length);
                }

                string Body = "";

                if (ToEmail == "NONE")
                {
                    ToEmail = System.Configuration.ConfigurationManager.AppSettings["BadEmailTo"];
                    Subject = "Buyback Order Header table ID " + ID;
                    Body = "Test PDF " + ID + ". ";
                }

                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["LIVEDB"].ToString());
                SqlDataAdapter da = new SqlDataAdapter();
                SqlCommand cmd = new SqlCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "PP_SendEmailwithAttachment_Hudson";
                SqlParameterCollection sqlParams = cmd.Parameters;
                sqlParams.AddWithValue("@userEmail", ToEmail);
                sqlParams.AddWithValue("@profileName", System.Configuration.ConfigurationManager.AppSettings["EmailSendingProfileName"]);
                sqlParams.AddWithValue("@subject", Subject);
                sqlParams.AddWithValue("@body", Body);
                sqlParams.AddWithValue("@attachment", filePath);
                sqlParams.AddWithValue("@BBCEmail", System.Configuration.ConfigurationManager.AppSettings["BBCEmails"]);
                sqlParams.AddWithValue("@fromEmails", System.Configuration.ConfigurationManager.AppSettings["EmailsFrom"]);
                sqlParams.AddWithValue("@orderID", System.Configuration.ConfigurationManager.AppSettings["EmailsFrom"]);
                sqlParams.AddWithValue("@Confirmation", System.Configuration.ConfigurationManager.AppSettings["EmailsFrom"]);
                sqlParams.AddWithValue("@ID", ID);
                conn.Open();
                cmd.Connection = conn;
                cmd.ExecuteReader();
                cmd.Dispose();
                conn.Close();
                conn.Dispose();

                Context.Response.Write("done");
            }
            catch (Exception e)
            {

                Context.Response.Write($"Error : {e.ToString()}");
            }            
        }
    }
}
