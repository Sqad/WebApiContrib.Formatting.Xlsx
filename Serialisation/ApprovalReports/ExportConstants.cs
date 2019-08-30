using System.Drawing;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports
{
    internal static class ExportConstants
    {
        internal const string ApprovalReportSheetName = "Approval Report",
                              EvenGroupColumnName= "IsEvenGroup",
                              IntExcelFormatTemplate = "#",
                              DateExcelFormatTemplate = "m/d/yyyy h:mm",
                              AccountingExcelFormatTemplate = "_($* #,##0.00_);_($* (#,##0.00);_($* - ??_);_(@_)";

        internal static Color EvenGroupDefaultColor = Color.LightGray;
    }
}
