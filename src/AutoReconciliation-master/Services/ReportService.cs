using System;
using System.IO;

namespace AutoReconciliation.Services
{
    class ReportService
    {
        public void ReportError(string cardCode, string linkedCardCode, Exception e, string xml)
        {
            if (linkedCardCode == "")
            {
                string report = $"Report on Error in AutoReconciliation Addon on {DateTime.Now.ToString("MM/dd/yyyy HH/mm")}\n\n";
                report += $"Business Partner CardCode -> [{cardCode}]\n\n";
                report += $"Error details:\n";
                report += e.ToString();
                report += "\n\n";
                report += "XML of Open Transactions:\n";
                report += xml;
                File.WriteAllText($"Logs/Log{DateTime.Now.ToString("MM/dd/yyyy HH/mm")}_{cardCode}.txt", report);
            }
            else
            {
                string report = $"Report on Error (Linked Account) in AutoReconciliation Addon on {DateTime.Now.ToString("MM/dd/yyyy HH/mm")}\n\n";
                report += $"Business Partner CardCode -> [{cardCode}]\n";
                report += $"Linked Business Partner CardCode -> [{linkedCardCode}]\n\n";
                report += $"Error details:\n";
                report += e.ToString();
                report += "\n\n";
                report += "XML of Open Transactions:\n";
                report += xml;
                File.WriteAllText($"Logs/Log{DateTime.Now.ToString("MM/dd/yyyy HH/mm")}_{cardCode}.txt", report);
            }
        }
    }
}
