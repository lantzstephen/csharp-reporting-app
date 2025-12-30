using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace Portfolio.Tools.Reporter
{
    /// <summary>
    /// HTML report generator with email and file output.
    /// Demonstrates: threading, reflection, LINQ, HTML templating, DataTable manipulation.
    /// </summary>
    static class Reporter
    {
        #region Format strings

        const string EmailSubjectFmt = "{0} {1}";
        const String formatDisplay = "M/d/yyyy h:mm:ss tt";
        const String formatTimeStamp = "yyyyMMdd_HHmmss";

        const string HtmlHeaderFormatEmail = @"
            <table style='font-family: Calibri, sans-serif; width: 100%'; padding: 0px;>
                <tr>
                    <span style='font-size: 11.0pt; color: Black;'>{0}</span><br><br>
                    <span style='font-size: 16.0pt; font-weight: bold; color: DarkBlue;'>{1}</span><br>
                    <span style='font-size: 11.0pt; font-weight: bold; color: FireBrick;'>{2}</span><br>
                    <span style='font-size: 11.0pt; font-weight: bold; color: DimGray;'>{3}</span>
                </tr>
            </table>";

        const string HtmlHeaderFormat = @"
            <table style='font-family: Calibri, sans-serif; font-size: 11.0pt; width: 100%; padding: 0px;'>
                <tr>
                    <td valign='bottom' align='center'>
                        <span style='font-size: 28.0pt; font-weight: bold; color: DarkBlue;'>{0}</span><br>
                        <span style='font-size: 11.0pt; font-weight: bold; color: FireBrick;'>{2}</span><br>
                        <span style='font-size: 11.0pt; font-weight: bold; color: DimGray;'>{3}</span>
                    </td>
                    <td width='700px' align='center'>
                        <img alt='Company Logo' src='{1}' />
                    </td>
                </tr>
            </table>
            <table style='font-family: Calibri, sans-serif; font-size: 11.0pt; width: 100%'>
                <tr>
                    <td colspan='3'><hr style='height: 2px; background-color: black;'/></td>
                </tr>
                <br><br>
            </table>";

        const string HtmlBodyFmt =
        @"  <html>
            <head>
                <style>
                    td.rptCell          {{ border: 1px solid black; border-top: none; color: black; font-size: 10.0pt; white-space: nowrap; padding: 7px;}}
                    td.rptCellHeader    {{ border: 1px solid black; color: DarkBlue; font-size: 11.0pt; text-align: center; white-space: nowrap; padding: 7px;}}
                    tr.rptRowHeader     {{ font-weight: bold; color: DarkBlue; background-color: lightgrey;}}
                    tr.rptRowAlt        {{ background-color: whitesmoke;}}
                </style>
            </head>
            <body style='font-family: Calibri, sans-serif; font-size: 11.0pt;'>
                {0}
                {1}
                <table style='border-collapse: collapse; width: 100%;'>
                    <br>
                    <tr><hr style='height: 2px; background-color: black;'/></tr>
                    <tr style='font-family: Calibri, sans-serif; font-size: 11.0pt;'>
                        <span>{2}</span>
                    </tr>
                </table>
            </body>
        </html>";

        const string tableFmt = @"
        <table style='border-collapse: collapse; width: 100%;'>
            <br>
            <tr style='font-family: Calibri, sans-serif; font-size: 11.0pt;'>
                <td width='1px' valign='baseline'>
                    <span style='color: DarkBlue; font-size: 14.0pt; font-style: italic; font-weight: bold;'>{1}&nbsp;&nbsp;</span>
                </td>
                <td valign='baseline'>
                    <span style='color: FireBrick;'>FILTER BY:</span>
                    <span style='color: DimGray;'>{4}&nbsp;&nbsp;&nbsp;</span>
                    <span style='color: FireBrick;'>SORT BY:</span>
                    <span style='color: DimGray;'>{3}&nbsp;&nbsp;&nbsp;</span>
                    <span style='color: FireBrick;'>ROWCOUNT:</span>
                    <span style='color: DimGray;'>{2}</span><br>
                </td>
            </tr>
        </table>
        <table style='border-collapse: collapse; width: 1024px;'>
            {5}
            {6}
        </table>
        {7}";

        #endregion

        static private WaitDialog wait;

        static private void Wait()
        {
            wait = new WaitDialog();
            Application.Run(wait);
        }

        /// <summary>
        /// Generate an HTML report from stored procedure results with optional grouping and sorting.
        /// </summary>
        static public void GenerateReport(
              DataConnection Connection
            , Report Meta
            , string Title = ""
            , string Header = ""
            , string Variables = ""
            , string Command = ""
            , string Parameters = ""
            , string GroupBy = ""
            , string GroupByList = ""
            , string SortBy = ""
            , string Filter = ""
            , string Recipients = ""
            , string OutputPath = ""
            , bool UseViewer = false)
        {
            ReportViewer rv = new ReportViewer();
            List<SqlParameter> variables;
            List<SqlParameter> parameters;

            string html = "";
            string headerHtml = "";
            string headerHtmlEmail = "";
            string tableHtml = "";
            string tableHtmlSummary = "";
            string tableHtmlDetail = "";
            string fileName = "";
            string outputPathFull = "";
            string outputPathHtml = "";

            try
            {
                // Show wait dialog on separate STA thread
                if (UseViewer)
                {
                    Thread t = new Thread(Wait);
                    t.SetApartmentState(ApartmentState.STA);
                    t.Start();
                }

                string logo = DataAccess.GetMasterSettingValue(Connection.MetaServer, "Common", "reports_logo");
                string reportDateTime = SystemLibrary.CurrentDatetime();
                string parameterListFinal = "";

                // Normalize properties
                Console.WriteLine("Processing report properties...");
                Title = ProcessProperty(Title);
                Header = ProcessProperty(Header);
                Command = ProcessProperty(Command);
                Filter = ProcessProperty(Filter);
                GroupBy = ProcessProperty(GroupBy, alterLetterCase: LetterCase.Upper);
                GroupByList = ProcessProperty(GroupByList, alterLetterCase: LetterCase.Upper);
                SortBy = ProcessProperty(SortBy, alterLetterCase: LetterCase.Upper);
                Recipients = ProcessProperty(Recipients, alterLetterCase: LetterCase.Lower);
                Variables = ProcessProperty(Variables);
                Parameters = ProcessProperty(Parameters);
                Meta.Sender = Meta.Sender == "" ? "reports@example.com" : Meta.Sender;
                Meta.Variables = ProcessProperty(Meta.Variables);
                Meta.Parameters = ProcessProperty(Meta.Parameters);

                // Merge meta and console variables and command parameters
                Console.WriteLine("Processing variables and command parameters");
                variables = MergeKVPLists(GenerateKVPList(Meta.Variables, null), GenerateKVPList(Variables, null));
                parameters = MergeKVPLists(GenerateKVPList(Meta.Parameters, variables), GenerateKVPList(Parameters, variables));

                foreach(SqlParameter p in parameters)
                {
                    string pval = p.Value.ToString().Replace(",", ", ");
                    parameterListFinal += $"; {p.ParameterName}={pval}";
                }

                parameterListFinal = parameterListFinal.TrimStart(';');

                if (parameterListFinal.Length > 500)
                {
                    parameterListFinal = parameterListFinal.Substring(1, 500) + " [...]";
                }

                // Get report data
                Console.WriteLine("Retrieving report data");
                DataTable detail = DataAccess.ExecSQL(Connection.Server, Connection.Database, Command,
                    CommandType.StoredProcedure, parameters.ToArray(), null, Connection.Timeout);
                DataTable summary = new DataTable();

                if (detail.Rows.Count == 0)
                {
                    tableHtml = GetTableHTML(detail, "DETAIL");
                }
                else
                {
                    // Apply filter using DataView
                    Console.WriteLine("Processing report filter");
                    if (Filter != "")
                    {
                        DataView dataView = new DataView(detail);
                        dataView.RowFilter = Filter;
                        detail = dataView.ToTable();
                    }

                    // Apply sort using DataView
                    if (SortBy != "")
                    {
                        DataView dv = new DataView(detail);
                        dv.Sort = SortBy;
                        detail = dv.ToTable();
                    }

                    // Generate summary section with grouping
                    Console.WriteLine("Processing report summary section");
                    if (GroupBy != "")
                    {
                        string groupByListColumnFinal = "";

                        if (GroupByList != "")
                        {
                            groupByListColumnFinal = "ASSOCIATED " + GroupByList + "(s)";
                        }

                        string[] groupBy = (GroupBy + ",").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        string countColumn = groupBy[0];

                        DataView dvGroup = new DataView(detail);
                        DataTable dtGroup = dvGroup.ToTable(true, groupBy);
                        dtGroup.Columns.Add("Count", typeof(int));

                        if (groupByListColumnFinal != "")
                        {
                            dtGroup.Columns.Add(groupByListColumnFinal, typeof(string));
                        }

                        foreach (DataRow dr in dtGroup.Rows)
                        {
                            dr["Count"] = detail.Compute("Count(" + countColumn + ")",
                                countColumn + " = '" + dr[countColumn] + "'");

                            if (groupByListColumnFinal != "")
                            {
                                List<string> valList = new List<string>();
                                foreach (DataRow r in detail.AsEnumerable()
                                    .Where(row => row.Field<String>(groupBy[0]) == (String)dr[groupBy[0]]))
                                {
                                    string val = (string)r[GroupByList];
                                    if (!valList.Contains(val)) valList.Add(val);
                                }

                                valList.Sort();
                                dr[groupByListColumnFinal] = String.Join(",", valList.ToArray());
                            }
                        }

                        DataView dvSort = new DataView(dtGroup);
                        dvSort.Sort = GroupBy;
                        summary = dvSort.ToTable();

                        tableHtmlSummary = GetTableHTML(summary, "SUMMARY", GroupBy, Filter, parameterListFinal);
                    }

                    // Generate detail section
                    Console.WriteLine("Processing report detail section");
                    tableHtmlDetail = GetTableHTML(detail, "DETAIL", SortBy, Filter, parameterListFinal);
                    tableHtml = tableHtmlSummary + tableHtmlDetail;
                }

                // Process report header
                Console.WriteLine("Processing report header text");
                string databaseFQName = $"{Connection.Server}.{Connection.Database}";
                string headerEmailIntro = (tableHtmlSummary != "" ? "Please review the following Report Summary.  " : "")
                    + $"See {(UseViewer ? "attached" : "link below")} for detailed report.";
                Header = Header == "" ? $"With data loaded through {reportDateTime}" : ProcessVariables(Header, variables);

                headerHtml = string.Format(HtmlHeaderFormat, Title, logo, databaseFQName, Header);
                headerHtmlEmail = string.Format(HtmlHeaderFormatEmail, headerEmailIntro, Title, databaseFQName, Header);

                // Generate filename
                Console.WriteLine("Processing file name");
                fileName = $"{Title.Replace(" -", "-").Replace("- ", "-").Replace(" ", "_")}_{DateTime.Now.ToString(formatTimeStamp)}";

                if (!UseViewer && OutputPath != "")
                {
                    if (!Directory.Exists(OutputPath))
                    {
                        Directory.CreateDirectory(OutputPath);
                    }

                    outputPathFull = $"{OutputPath.TrimEnd(Path.DirectorySeparatorChar)}\\{fileName}.htm";
                    outputPathHtml = $"This report is accessible at the following address: <a href='{outputPathFull}'>{outputPathFull}</a>";
                }

                // Final html creation
                Console.WriteLine("Generating final report HTML");
                html = string.Format(HtmlBodyFmt, headerHtml, tableHtml, outputPathHtml);

                Console.WriteLine("Preparing email HTML content");
                string emailBody = string.Format(HtmlBodyFmt, headerHtmlEmail,
                    tableHtmlSummary != "" ? tableHtmlSummary : tableHtml, outputPathHtml);
                string emailSubject = string.Format(EmailSubjectFmt, Connection.Database, Title);

                if (UseViewer)
                {
                    rv.Html = html;
                    rv.FileName = fileName;
                    rv.Summary = summary;
                    rv.Detail = detail;
                    rv.EmailBody = emailBody;
                    rv.EmailSubject = emailSubject;
                    rv.EmailRecipients = Recipients;
                    rv.EmailSender = Meta.Sender;
                    rv.Titlebar = $"Report Viewer v{Application.ProductVersion} - {System.Environment.UserName.ToUpper()} Connected to {Connection.MetaServer}";
                    if (wait != null) wait.Close();
                    rv.ShowDialog();
                }
                else
                {
                    DataAccess.WriteFile(outputPathFull, html);
                    DataAccess.SendEmail(emailBody, emailSubject, Recipients, Meta.Sender);
                }
            }
            catch (Exception e)
            {
                if (wait != null)
                    wait.Close();

                if (rv != null)
                    rv.Close();

                throw e;
            }
        }

        /// <summary>
        /// Merge two lists of SqlParameters, avoiding duplicates.
        /// </summary>
        private static List<SqlParameter> MergeKVPLists(List<SqlParameter> sourceKVPList, List<SqlParameter> targetKVPList)
        {
            foreach (SqlParameter s in sourceKVPList)
            {
                bool inTarget = false;

                foreach (SqlParameter t in targetKVPList)
                {
                    if (t.ParameterName.Equals(s.ParameterName, StringComparison.CurrentCultureIgnoreCase))
                        inTarget = true;
                }

                if (!inTarget)
                {
                    targetKVPList.Add(new SqlParameter() {
                        ParameterName = s.ParameterName,
                        Value = s.Value,
                        SqlDbType = s.SqlDbType
                    });
                }
            }

            return targetKVPList;
        }

        /// <summary>
        /// Parse comma-separated key=value pairs into SqlParameters.
        /// </summary>
        private static List<SqlParameter> GenerateKVPList(string csvText, List<SqlParameter> variables = null)
        {
            List<SqlParameter> KVPList = new List<SqlParameter>();

            if (csvText != "")
            {
                csvText = csvText + ",";
                string[] propertyArray = csvText.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string property in propertyArray)
                {
                    int i = Array.IndexOf(propertyArray, property);

                    if (variables != null)
                        propertyArray[i] = ProcessVariables(propertyArray[i], variables);

                    string[] e = propertyArray[i].Split('=');
                    e[0] = e[0].Trim().Trim(',').Replace("@", "");
                    e[1] = e[1].Trim().Trim(',');

                    KVPList.Add(new SqlParameter() {
                        ParameterName = e[0],
                        Value = e[1],
                        SqlDbType = SqlDbType.VarChar
                    });
                }
            }

            return KVPList;
        }

        /// <summary>
        /// Normalize a property value with optional case conversion.
        /// </summary>
        private static string ProcessProperty(string value, string valueWhenEmpty = "",
            LetterCase alterLetterCase = LetterCase.NoChange,
            bool removeDoubleSpace = true, bool removeCommaSpace = true, bool removeTrailingComma = true)
        {
            if (!string.IsNullOrEmpty(value))
                value = value.Trim();

            if (!string.IsNullOrEmpty(valueWhenEmpty))
                valueWhenEmpty = valueWhenEmpty.Trim();

            value = string.IsNullOrEmpty(value) ? valueWhenEmpty : value;

            if (!string.IsNullOrEmpty(value))
            {
                if (removeDoubleSpace) value = value.Replace("  ", " ");
                if (removeCommaSpace) value = value.Replace(" ,", ",").Replace(", ", ",");
                if (removeTrailingComma) value = value.Trim(',');

                switch (alterLetterCase)
                {
                    case LetterCase.Lower:
                        value = value.ToLower();
                        break;

                    case LetterCase.Upper:
                        value = value.ToUpper();
                        break;
                }
            }

            return value;
        }

        /// <summary>
        /// Replace variable placeholders with values and execute system functions.
        /// </summary>
        private static string ProcessVariables(string property, List<SqlParameter> variables)
        {
            if (variables.Count > 0)
            {
                foreach (SqlParameter v in variables)
                {
                    property = Regex.Replace(property, Regex.Escape($"[{v.ParameterName}]"),
                        v.Value.ToString(), RegexOptions.IgnoreCase);
                }
            }

            property = ExecuteSystemFunctions(property);

            return property;
        }

        /// <summary>
        /// Generate HTML table from DataTable.
        /// </summary>
        static string GetTableHTML(DataTable data, string tableTitle, string sortBy = "",
            string filterBy = "", string parameters = "")
        {
            string headerRowHtml = "";
            string rowsHtml = "";
            string footer = "";
            int rowCount = data.Rows.Count;
            StringBuilder builder = new StringBuilder();

            // Build header row
            builder.AppendLine(@"<tr class=""rptRowHeader"">");
            for (int i = 0; i < data.Columns.Count; i++)
            {
                builder.AppendLine($@"<td class=""rptCellHeader"">{data.Columns[i].ColumnName.ToUpper()}</td>");
            }
            builder.AppendLine(@"</tr>");
            headerRowHtml = builder.ToString();

            builder = new StringBuilder();

            // Build data rows with alternating colors
            for (int i = 0; i < data.Rows.Count; i++)
            {
                DataRow dr = data.Rows[i];
                builder.AppendLine((i % 2) != 0 ? @"<tr class=""rptRowAlt"">" : @"<tr>");

                for (int j = 0; j < data.Columns.Count; j++)
                {
                    string val = dr[j].ToString();
                    val = val.Length > 136 ? val.Substring(0, 136) + "..." : val;
                    builder.Append(@"<td class=""rptCell"" ");
                    builder.Append(@">" + val + "</td>");
                }

                builder.AppendLine(@"</tr>");
            }

            rowsHtml = builder.ToString();

            if (rowCount == 0)
                footer = @"<span style='font-family: Calibri, sans-serif; font-size: 11.0pt; color: DimGray; font-weight: bold;'>No data is available for the report as currently defined.</span>";

            sortBy = sortBy == "" ? "None" : sortBy.Replace(",", ", ");
            filterBy = filterBy == "" ? "None" : filterBy;

            return string.Format(tableFmt, data.Columns.Count - 1, tableTitle, rowCount.ToString(),
                sortBy, filterBy, headerRowHtml, rowsHtml, footer);
        }

        /// <summary>
        /// Execute embedded system function calls using reflection.
        /// Demonstrates: reflection, MethodInfo.Invoke(), dynamic method invocation.
        /// </summary>
        static string ExecuteSystemFunctions(string text)
        {
            int funcStart = 0;
            int funcEnd = 0;
            string sysCall = "";
            string sysCallRet = "";
            string returnVal = text;

            while ((funcStart = returnVal.IndexOf("{", funcStart)) != -1)
            {
                funcStart++;
                funcEnd = returnVal.IndexOf("}", funcStart);
                sysCall = returnVal.Substring(funcStart, funcEnd - funcStart);
                sysCallRet = CallSystemFunction(sysCall);

                if (sysCallRet == "")
                {
                    throw new Exception($"System function call {sysCall} unexpectedly returned no result.");
                }

                returnVal = $"{returnVal.Substring(0, funcStart - 1)}{sysCallRet}{returnVal.Substring(funcEnd + 1)}";
            }

            return returnVal;
        }

        static private string ExtractText(string searchText, string startMarker = "",
            string endMarker = "", int startPosition = 0)
        {
            string returnValue = "";

            int textStart = startMarker == "" ? 0 : searchText.IndexOf(startMarker, startPosition) + 1;
            int textEnd = endMarker == "" ? searchText.Length - 1 : searchText.IndexOf(endMarker, textStart);

            if (textStart > -1 && textEnd > -1 && textStart < textEnd)
            {
                returnValue = searchText.Substring(textStart, textEnd - textStart);
            }

            return returnValue;
        }

        /// <summary>
        /// Invoke a method by name using reflection.
        /// </summary>
        static private string CallSystemFunction(string functionCall)
        {
            string returnValue = "";

            string function = ExtractText(functionCall, "", "(");
            string parameters = ExtractText(functionCall, "(", ")");

            Type type = typeof(SystemLibrary);
            MethodInfo methodInfo = type.GetMethod(function);

            if (!string.IsNullOrEmpty(parameters))
            {
                object[] parametersArray = Array.ConvertAll(parameters.Split(','), p => p.Trim());
                returnValue = (string)methodInfo.Invoke(null, parametersArray);
            }
            else
            {
                returnValue = (string)methodInfo.Invoke(null, null);
            }

            return returnValue;
        }
    }
}
