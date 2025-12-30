using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;
using System.IO;
using System.Net.Mail;
using System.Collections.Generic;
using System.Text;
using System.Net.Mime;
using OfficeOpenXml;
using System.Configuration;

namespace Portfolio.Tools.Reporter
{
    /// <summary>
    /// Data connection configuration.
    /// </summary>
    internal class DataConnection
    {
        public string MetaServer { get; set; }
        public string Server { get; set; }
        public string Database { get; set; }
        public string Timeout { get; set; }
    }

    /// <summary>
    /// Report configuration from metadata.
    /// </summary>
    internal class Report
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public string Header { get; set; }
        public string Variables { get; set; }
        public string Command { get; set; }
        public string Parameters { get; set; }
        public string GroupBy { get; set; }
        public string GroupByList { get; set; }
        public string SortBy { get; set; }
        public string Filter { get; set; }
        public string Recipients { get; set; }
        public string Sender { get; set; }
        public string OutputPath { get; set; }
        public string UseViewer { get; set; }
    }

    internal enum LetterCase
    {
        Upper,
        Lower,
        NoChange
    }

    /// <summary>
    /// Data access layer for report generation.
    /// Demonstrates: IDisposable patterns, parameterized queries, Excel generation, email with attachments.
    /// </summary>
    static class DataAccess
    {
        const string cnTemplate = "Data Source={0}; Initial Catalog={1}; Trusted_Connection=True; Pooling=False; Connect Timeout={2}";

        public static string Environment
        {
            get
            {
                string value = "";
                try
                {
                    value = ConfigurationManager.AppSettings["Environment"];
                }
                catch { }
                return value;
            }
        }

        public static bool SQLDatabaseExists(string Server, string Database)
        {
            bool result = false;
            try
            {
                ExecSQL(Server, Database, "SELECT 1", CommandType.Text, null, null, "1");
                result = true;
            }
            catch { }
            return result;
        }

        public static void WriteFile(string FullPath, string Text)
        {
            if (FullPath != "")
            {
                Console.WriteLine("Saving report to output path");
                using (var w = new StreamWriter(FullPath))
                {
                    w.WriteLine(Text);
                    w.Flush();
                }
            }
        }

        public static void WriteExcel(string FullPath, DataSet Data)
        {
            ExcelPackage spreadSheet = ExcelFile(Data);
            FileInfo file = new FileInfo(FullPath);
            spreadSheet.SaveAs(file);
        }

        /// <summary>
        /// Generate Excel file from DataSet using EPPlus.
        /// Demonstrates: working with third-party libraries, DataTable iteration.
        /// </summary>
        public static ExcelPackage ExcelFile(DataSet Data)
        {
            ExcelPackage spreadSheet = new ExcelPackage();

            foreach(DataTable t in Data.Tables)
            {
                if (t.Rows.Count > 0)
                {
                    var worksheet = spreadSheet.Workbook.Worksheets.Add(t.TableName);
                    worksheet.Cells["A1"].LoadFromDataTable(t, true);

                    // Format date columns
                    for (int colIndex = 0; colIndex <= t.Columns.Count - 1; colIndex++)
                    {
                        string type = t.Columns[colIndex].DataType.Name;

                        if (type == "DateTime")
                        {
                            worksheet.Column(colIndex + 1).Style.Numberformat.Format = "mm/dd/yyyy hh:mm:ss AM/PM";
                        }
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }
            }

            return spreadSheet;
        }

        /// <summary>
        /// Send HTML email with optional Excel or text attachment.
        /// Demonstrates: SMTP, MemoryStream, attachment handling.
        /// </summary>
        public static void SendEmail(
              string HtmlBody
            , string Subject
            , string Recipients
            , string AttachmentName = ""
            , string AttachmentText = ""
            , DataSet AttachmentDataSet = null
            , string Sender = "reports@example.com")
        {
            if (Recipients == "")
                return;

            SmtpClient client = new SmtpClient("SMTP");
            MailMessage mailMessage = new MailMessage();
            mailMessage.Subject = Subject;
            mailMessage.From = new MailAddress(Sender);
            mailMessage.IsBodyHtml = true;

            Recipients = Recipients.Replace(";", ",") + ",";

            foreach (string recipient in Recipients.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                mailMessage.To.Add(new MailAddress(recipient.Trim()));
            }

            AlternateView view = AlternateView.CreateAlternateViewFromString(HtmlBody, null, "text/html");

            foreach (var linkedResource in new List<LinkedResource>())
            {
                view.LinkedResources.Add(linkedResource);
            }

            mailMessage.AlternateViews.Add(view);

            using (MemoryStream memoryStream = new MemoryStream())
            {
                if (AttachmentName != "" && (AttachmentText != "" || AttachmentDataSet != null))
                {
                    if (AttachmentDataSet != null)
                    {
                        ExcelPackage contentAsExcel = ExcelFile(AttachmentDataSet);
                        contentAsExcel.SaveAs(memoryStream);
                    }
                    else if (AttachmentText != "")
                    {
                        byte[] contentAsBytes = Encoding.UTF8.GetBytes(AttachmentText);
                        memoryStream.Write(contentAsBytes, 0, contentAsBytes.Length);
                    }

                    memoryStream.Seek(0, SeekOrigin.Begin);

                    ContentType contentType = new ContentType();
                    contentType.MediaType = MediaTypeNames.Application.Octet;
                    contentType.Name = AttachmentName;
                    Attachment attachment = new Attachment(memoryStream, contentType);

                    mailMessage.Attachments.Add(attachment);
                }

                Console.WriteLine($"Sending report to recipient(s) via account '{Sender}'");
                client.Send(mailMessage);
            }
        }

        public static bool SQLRoutineExists(string Server, string Database, string Routine)
        {
            string routine = null;
            string sql =
                $"SELECT routine = CAST(o.name AS varchar(MAX)) " +
                $"FROM sys.objects o " +
                $"JOIN sys.schemas s " +
                $"  ON s.schema_id = o.schema_id " +
                $"  AND s.name IN ('dbo', 'rpt') " +
                $"WHERE object_id = OBJECT_ID('{Routine}') AND type IN ('P', 'SN')";

            DataTable dt = ExecSQL(Server, Database, sql, CommandType.Text, null);
            if (dt.Rows.Count == 1) routine = dt.Rows[0].Field<string>("routine");

            return (!string.IsNullOrEmpty(routine));
        }

        /// <summary>
        /// Get database connection from XML configuration.
        /// Demonstrates: XML parsing with LINQ to XML.
        /// </summary>
        public static DataConnection GetConnection(string MetaServer, string Database)
        {
            DataConnection returnValue = new DataConnection();
            returnValue.MetaServer = MetaServer.ToUpper();

            string DatabaseValidated = "";
            string cloneConnectionInfoXml = GetConfigValue(returnValue.MetaServer, "10", "CloneConnectionInfoXml");

            if (cloneConnectionInfoXml != "")
            {
                XDocument xdoc = XDocument.Parse(cloneConnectionInfoXml);
                var clones = xdoc.Descendants("Clone");

                foreach (var clone in clones)
                {
                    if (SystemLibrary.IsMatch(clone.Attribute("name").Value.ToString(), Database))
                    {
                        returnValue.Server = clone.Attribute("server").Value.ToString();
                        returnValue.Timeout = clone.Attribute("timeout").Value.ToString();
                        DatabaseValidated = LookupString(returnValue.Server, "master", "sys.databases",
                            "name", $"name = '{Database}'", "");

                        if (DatabaseValidated == "")
                        {
                            throw new Exception($"Specified database {Database} does not exist on server {returnValue.Server}.");
                        }
                        {
                            returnValue.Database = DatabaseValidated;
                        }
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Get report configuration from XML metadata.
        /// </summary>
        public static Report GetReport(string MetaServer, string Name)
        {
            Report returnValue = new Report();
            returnValue.Name = Name;

            string reportPresetInfoXml = GetConfigValue(MetaServer, "10", "ReportPresetInfoXml");
            XDocument xdoc = XDocument.Parse(reportPresetInfoXml);
            var reports = xdoc.Descendants("Report");

            foreach (var report in reports)
            {
                if (SystemLibrary.IsMatch(report.Attribute("name").Value.ToString(), Name))
                {
                    returnValue.Title = report.Attribute("title")?.Value.ToString() ?? Name;
                    returnValue.Header = report.Attribute("header")?.Value.ToString() ?? "";
                    returnValue.Variables = report.Attribute("variables")?.Value.ToString() ?? "";
                    returnValue.Command = report.Attribute("command")?.Value.ToString().Replace("[","").Replace("]", "") ?? "";
                    returnValue.Parameters = report.Attribute("parameters")?.Value.ToString() ?? "";
                    returnValue.GroupBy = report.Attribute("groupby")?.Value.ToString() ?? "";
                    returnValue.GroupByList = report.Attribute("groupbylist")?.Value.ToString() ?? "";
                    returnValue.SortBy = report.Attribute("sortby")?.Value.ToString() ?? "";
                    returnValue.Filter = report.Attribute("filter")?.Value.ToString() ?? "";
                    returnValue.Recipients = report.Attribute("recipients")?.Value.ToString() ?? "";
                    returnValue.Sender = report.Attribute("sender")?.Value.ToString() ?? "";
                    returnValue.OutputPath = report.Attribute("outputpath")?.Value.ToString() ?? "";
                    returnValue.UseViewer = report.Attribute("useviewer")?.Value.ToString() ?? "0";
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Execute SQL query with support for table-valued parameters.
        /// Demonstrates: IDisposable, SqlParameter arrays, structured parameters.
        /// </summary>
        public static DataTable ExecSQL(string Server, string Database, string CommandText,
            CommandType CommandType, SqlParameter[] StaticParameters = null,
            DataTable DynamicParameters = null, string Timeout = "15")
        {
            DataTable results = new DataTable();
            string connString = string.Format(cnTemplate, Server, Database, Timeout);

            using (SqlConnection conn = new SqlConnection(string.Format(cnTemplate, Server, Database, Timeout)))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandType = CommandType;

                    if (StaticParameters != null)
                    {
                        cmd.Parameters.AddRange(StaticParameters);
                    }

                    if (DynamicParameters != null)
                    {
                        SqlParameter parameter;
                        parameter = cmd.Parameters.AddWithValue("@DynamicParameters", DynamicParameters);
                        parameter.SqlDbType = SqlDbType.Structured;
                        parameter.TypeName = "dbo.KVP";
                    }

                    cmd.CommandText = CommandText;
                    cmd.Connection = conn;
                    cmd.CommandTimeout = 0;
                    conn.Open();

                    using (IDataReader r = cmd.ExecuteReader())
                    {
                        results.Load(r);
                    }

                    if (StaticParameters != null)
                    {
                        cmd.Parameters.Clear();
                    }
                }
            }

            return results;
        }

        public static string LookupString(string Server, string Database, string TableName,
            string ColumnName, string Filter, string DefaultIfEmpty = null)
        {
            string value = null;
            string sql =
                $"SELECT TOP 1 value = CAST({ColumnName} as varchar(MAX)) " +
                $"FROM {TableName} " +
                $"WHERE {Filter}";

            DataTable dt = ExecSQL(Server, Database, sql, CommandType.Text, null);
            if (dt.Rows.Count == 1) value = dt.Rows[0].Field<string>("value");
            if (value == null) value = DefaultIfEmpty;

            value = value.Replace("{ENV}", Environment);

            return value;
        }

        public static int? LookupInt(string Server, string Database, string TableName,
            string ColumnName, string Filter, int? DefaultIfEmpty = null)
        {
            int? value = null;

            DataTable dt = ExecSQL(Server, Database,
                $"SELECT TOP 1 value = CAST({ColumnName} as varchar(MAX)) FROM {TableName} WHERE {Filter}",
                CommandType.Text, null);
            if (dt.Rows.Count == 1) value = dt.Rows[0].Field<int?>("value");
            if (value == null) value = DefaultIfEmpty;

            return value;
        }

        public static DateTime? LookupDateTime(string Server, string Database, string TableName,
            string ColumnName, string Filter, DateTime? DefaultIfEmpty = null)
        {
            DateTime? value = null;

            DataTable dt = ExecSQL(Server, Database,
                $"SELECT TOP 1 value = CAST({ColumnName} as varchar(MAX)) FROM {TableName} WHERE {Filter}",
                CommandType.Text, null);
            if (dt.Rows.Count == 1) value = dt.Rows[0].Field<DateTime?>("value");
            if (value == null) value = DefaultIfEmpty;

            return value;
        }

        public static string GetConfigValue(string MetaServer, string ApplicationId, string ConfigName)
        {
            return LookupString(MetaServer, "Common", "Utility.InterfaceConfig", "ConfigValue",
                $"ApplicationId = {ApplicationId} AND ConfigName = '{ConfigName}'", "");
        }

        public static string GetMasterSettingValue(string Server, string Database, string settingName)
        {
            return LookupString(Server, Database, "dbo.fnGetMaster_settings_def(NULL)", "setting_value",
                $"setting_name = '{settingName}'");
        }
    }
}
