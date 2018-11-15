using System;
using System.Data;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views
{
    public class UnformattedExcelDataRow : ExcelDataRow
    {
        public UnformattedExcelDataRow(DataRow dataRow)
            : base(dataRow)
        {
        }

        protected override object ParseValue(string columnName)
        {
            var data = this[columnName];
            switch (columnName)
            {
                case "[Startdate]":
                case "[Enddate]":
                    return ((DateTime)data).ToString("dd-MM-yyyy");
                default:
                    return data;
            }
        }
    }
}
