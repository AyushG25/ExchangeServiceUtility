using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Text.RegularExpressions;

namespace ExchangeServiceUtility
{
    class TemplateCreator
    {
        public static string ConvertDatatableToHtml(DataTable table)
        {
            if (table == null || table.Rows.Count == 0)
            {
                return string.Empty;
            }

            StringBuilder htmlTable = new StringBuilder();

            htmlTable.Append("<table border='1' align='center'>");

            htmlTable.Append("<tr>");
            foreach (DataColumn col in table.Columns)
            {
                htmlTable.Append("<th>" + Convert.ToString(col.ColumnName) + "</th>");
            }
            htmlTable.Append("</tr>");

            //Add the data rows. 
            foreach (DataRow row in table.Rows)
            {
                htmlTable.Append("<tr>");
                foreach (DataColumn col in table.Columns)
                {
                    htmlTable.Append("<td>" + Convert.ToString(row[col.ColumnName]) + "</td>");
                }
                htmlTable.Append("</tr>");
            }

            htmlTable.Append("</table>");

            return htmlTable.ToString();
        }

        public static string CreateEmailTemplate<T>(T model)
        {
            string html = File.ReadAllText(ConfigurationManager.AppSettings["HTMLTemplatePath"]);
            string regex = @"Model\.\w+";
            var parameterList = Regex.Matches(html, regex);
            foreach (var parameter in parameterList)
            {
                string modelProperty = parameter.ToString().Replace("Model.", "");
                var m = model.GetType().GetProperty(modelProperty).GetValue(model);
                html = html.Replace(Convert.ToString(parameter), Convert.ToString(m));
            }
            return html;
        }
    }
}
