using System;
using System.Collections.Generic;
using System.Text;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports
{
    internal static class ExportConstants
    {
        internal const string ApprovalReportSheetName = "Approval Report",
                              IntExcelFormatTemplate = "#",
                              DateExcelFormatTemplate = "m/d/yyyy h:mm",
                              AccountingExcelFormatTemplate = "_($* #,##0.00_);_($* (#,##0.00);_($* - ??_);_(@_)";
    }
}
