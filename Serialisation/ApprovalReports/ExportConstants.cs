using System;
using System.Drawing;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports
{
    internal static class ExportConstants
    {
        internal const string ApprovalReportSheetName = "Approval Report",
                              BillingReportSheetName = "Billing Report",
                              EvenGroupColumnName= "IsEvenGroup",
                              CurrencySymbolColumnName = "CurrencySymbol",
                              IntExcelFormatTemplate = "#",
                              DateExcelFormatTemplate = "m/d/yyyy h:mm";

        internal static Func<string, string> CreateAccountingExcelFormatTemplate = currencySymbol => $"_({currencySymbol}* #,##0.00_);_({currencySymbol}* (#,##0.00);_({currencySymbol}* - ??_);_(@_)";

        internal static Color EvenGroupDefaultColor = Color.LightGray;
    }
}
