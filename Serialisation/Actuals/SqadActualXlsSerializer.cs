using SQAD.MTNext.Business.Models.Actual;
using SQAD.MTNext.Business.Models.Common.Enums;
using SQAD.XlsxExportImport.Base.Attributes;
using SQAD.XlsxExportImport.Base.Builders;
using SQAD.XlsxExportImport.Base.Interfaces;
using SQAD.XlsxExportImport.Base.Models;
using SQAD.XlsxExportImport.Base.Serialization;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Actuals
{
    class SqadActualXlsSerializer : IXlsxSerializer
    {
        public SerializerType SerializerType => SerializerType.Default;
        public bool CanSerializeType(Type valueType, Type itemType)
        {
            return valueType == typeof(ActualExport);
        }

        public void Serialize(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder sheetBuilderOverride)
        {
            if (!(value is ActualExport ae))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var export = (ActualExport)value;

            CreatePropertySheet(document, export);
            PopulateSheets(document, export);
            CreateReferenceSheet(document, export.Id);


        }
        private void CreateReferenceSheet(IXlsxDocumentBuilder document, int id)
        {
            var builder = new XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder("Reference",isPreservationSheet:true,isHidden:true);
            builder.ActualRow = true;

            document.AppendSheet(builder);

            var excelinfo = new ExcelColumnInfo(id.ToString(), null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(" ", null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo("Version", null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo("2.1", null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

        }

        private void CreatePropertySheet( IXlsxDocumentBuilder document, ActualExport value)
        {
            string name = "Properties";

            //var instructionsDataTable = tables[InstructionsTableName];

            var properties =  new XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder(name);
            properties.ActualRow = true;

            document.AppendSheet(properties);

            AppendColumnsAndRows(properties, value);
        }
        private void AppendColumnsAndRows(SqadXlsxSheetBuilderBase sheetBuilder, ActualExport value)
        {
            var exp = "Export Properties";
            var sec = " ";
            var excelinfo = new ExcelColumnInfo(exp, null, new ExcelColumnAttribute(), null);
            sheetBuilder.AppendColumnHeaderRowItem(excelinfo);
            excelinfo = new ExcelColumnInfo(sec, null, new ExcelColumnAttribute(), null);
            sheetBuilder.AppendColumnHeaderRowItem(excelinfo);


            List<ExcelCell> row = new List<ExcelCell>();
            ExcelCell cell = new ExcelCell();


            cell = new ExcelCell();
            cell.CellHeader = exp;
            cell.CellValue = $"Resource Set";
            row.Add(cell);

            cell = new ExcelCell();
            cell.CellHeader = sec;
            cell.CellValue = $"{value.ResourceName}";
            row.Add(cell);
            sheetBuilder.AppendRow(row);

            cell = new ExcelCell();
            cell.CellHeader = exp;
            cell.CellValue = $"Client";
            row.Add(cell);

            cell = new ExcelCell();
            cell.CellHeader = sec;
            cell.CellValue = $"{value.ClientName}";
            row.Add(cell);
            sheetBuilder.AppendRow(row);

            cell = new ExcelCell();
            cell.CellHeader = exp;
            cell.CellValue = $"From Date";
            row.Add(cell);

            cell = new ExcelCell();
            cell.CellHeader = sec;
            cell.CellValue = $"{value.Request.FromDate.ToString("dddd, MMMM dd, yyyy")}";
            row.Add(cell);
            sheetBuilder.AppendRow(row);

            cell = new ExcelCell();
            cell.CellHeader = exp;
            cell.CellValue = $"To Date";
            row.Add(cell);

            cell = new ExcelCell();
            cell.CellHeader = sec;
            cell.CellValue = $"{value.Request.ToDate.ToString("dddd, MMMM dd, yyyy")}";
            row.Add(cell);
            sheetBuilder.AppendRow(row);

            cell = new ExcelCell();
            cell.CellHeader = exp;
            cell.CellValue = $"Created By";
            row.Add(cell);

            cell = new ExcelCell();
            cell.CellHeader = sec;
            cell.CellValue = $"{value.CreatedBy}";
            row.Add(cell);
            sheetBuilder.AppendRow(row);

            cell = new ExcelCell();
            cell.CellHeader = exp;
            cell.CellValue = $"Created On";
            row.Add(cell);

            cell = new ExcelCell();
            cell.CellHeader = sec;
            cell.CellValue = $"{value.Request.CreateDate.ToString("dddd, MMMM dd, yyyy HH:mm")}";
            row.Add(cell);
            sheetBuilder.AppendRow(row);

            cell = new ExcelCell();
            cell.CellHeader = exp;
            cell.CellValue = string.Empty;
            row.Add(cell);

            cell = new ExcelCell();
            cell.CellHeader = sec;
            cell.CellValue = string.Empty;
            row.Add(cell);
            sheetBuilder.AppendRow(row);

            //PopulateNames(sheetBuilder, exp, sec, "Products", value.Flights.Select(x=>x.ProductName).Distinct().ToList());
            PopulateNames(sheetBuilder, exp, sec, "Products", value.ProductNames);
            if (value.PlanTypeNames.Any())
                PopulateNames(sheetBuilder, exp, sec, "Plan Types", value.PlanTypeNames);



        }

        private void PopulateNames(SqadXlsxSheetBuilderBase sheetBuilder, string exp, string sec, string cellname, List<string> names = null)
        {

            bool emptycell = false;
            foreach (var name in names)
            {
                List<ExcelCell> row = new List<ExcelCell>();
                var cell = new ExcelCell();
                cell.CellHeader = exp;
                cell.CellValue = (!emptycell) ? cellname : string.Empty;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = sec;
                cell.CellValue = $"{name}";
                row.Add(cell);
                sheetBuilder.AppendRow(row);
                emptycell = true;
            }

        }
        private void PopulatePlannedActualColumns(XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder builder, ActualWorksheet item, bool actual = false)
        {
            switch (item.MediaType.MediaClass)
            {
                case (int)MediaClass.Print:
                case (int)MediaClass.Newspaper:
                case (int)MediaClass.ReachBased:
                    if (!actual)
                    {
                        //visible
                        var excelinfo = new ExcelColumnInfo(item.PlannedProduction, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);
                        //unvisible
                        excelinfo = new ExcelColumnInfo(item.PlannedImpressions, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedClicks, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedLeads, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedRichMedia, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedAdServing, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);


                    }
                    else
                    {
                        var excelinfo = new ExcelColumnInfo(item.ActualProduction, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualImpressions, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualClicks, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualLeads, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualRichMedia, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualAdServing, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                    }
                    break;
                case (int)MediaClass.Interactive:
                    if (!actual)
                    {
                        var excelinfo = new ExcelColumnInfo(item.PlannedProduction, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedImpressions, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedClicks, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedLeads, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedRichMedia, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedAdServing, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);
                    }
                    else
                    {
                        var excelinfo = new ExcelColumnInfo(item.ActualProduction, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualImpressions, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualClicks, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualLeads, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualRichMedia, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualAdServing, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);
                    }
                    break;

                default:
                    if (!actual)
                    {
                        var excelinfo = new ExcelColumnInfo(item.PlannedProduction, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedImpressions, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedClicks, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedLeads, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedRichMedia, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.PlannedAdServing, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                    }
                    else
                    {
                        var excelinfo = new ExcelColumnInfo(item.ActualProduction, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualImpressions, null, new ExcelColumnAttribute(), null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualClicks, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualLeads, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualRichMedia, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                        excelinfo = new ExcelColumnInfo(item.ActualAdServing, null, new ExcelColumnAttribute() { IsHidden = true }, null);
                        builder.AppendColumnHeaderRowItem(excelinfo);

                    }
                    break;

            }

        }
        private void FormatWorkSheet(XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder builder, ActualWorksheet item, Dictionary<string, string> custom)
        {
            var excelinfo = new ExcelColumnInfo(item.Data, null, new ExcelColumnAttribute() { IsHidden = true }, null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.PlanName, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.VehicleName, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.ProductName, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.SubtypeName, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.UnitName, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.MarketName, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.DemoName, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            if (custom != null && custom.Count > 0)
            {
                foreach (var it in custom)
                {
                    excelinfo = new ExcelColumnInfo(it.Key, null, new ExcelColumnAttribute(), null);
                    builder.AppendColumnHeaderRowItem(excelinfo);
                }
            }
            else
            {
                excelinfo = new ExcelColumnInfo(item.CreativeName, null, new ExcelColumnAttribute(), null);
                builder.AppendColumnHeaderRowItem(excelinfo);

                excelinfo = new ExcelColumnInfo(item.FundingSourceName, null, new ExcelColumnAttribute(), null);
                builder.AppendColumnHeaderRowItem(excelinfo);
            }

            excelinfo = new ExcelColumnInfo(item.StartDate, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.EndDate, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            //planned
            excelinfo = new ExcelColumnInfo(item.PlannedGRPs, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.PlannedTRPs, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            var hidden = !(Convert.ToBoolean(item.MediaType.EnableWeeklyReach) || Convert.ToInt16(item.MediaType.ReachType) > 0);
            excelinfo = new ExcelColumnInfo(item.PlannedReach, null, new ExcelColumnAttribute() { IsHidden=hidden}, null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.PlannedFrequency, null, new ExcelColumnAttribute() { IsHidden = hidden }, null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            var shownet = Convert.ToBoolean(item.MediaType.ExternalActualsNet);
            excelinfo = new ExcelColumnInfo(item.PlannedGross, null, new ExcelColumnAttribute() {IsHidden = shownet }, null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.PlannedNet, null, new ExcelColumnAttribute() { IsHidden =!shownet}, null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            PopulatePlannedActualColumns(builder, item);

            //actual
            excelinfo = new ExcelColumnInfo(item.ActualGRPs, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.ActualTRPs, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.ActualReach, null, new ExcelColumnAttribute() { IsHidden = hidden }, null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.ActualFrequency, null, new ExcelColumnAttribute() { IsHidden = hidden }, null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            //if (!shownet)
            //{
            //    excelinfo = new ExcelColumnInfo(item.ActualGross, null, new ExcelColumnAttribute(), null);
            //    builder.AppendColumnHeaderRowItem(excelinfo);
            //}
            //else
            //{
            //    excelinfo = new ExcelColumnInfo(item.ActualNet, null, new ExcelColumnAttribute(), null);
            //    builder.AppendColumnHeaderRowItem(excelinfo);
            //}

            excelinfo = new ExcelColumnInfo(item.ActualGross, null, new ExcelColumnAttribute() { IsHidden = shownet }, null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.ActualNet, null, new ExcelColumnAttribute() { IsHidden = !shownet }, null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            PopulatePlannedActualColumns(builder, item, true);

            excelinfo = new ExcelColumnInfo(item.DateActualized, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);

            excelinfo = new ExcelColumnInfo(item.ActualizedBy, null, new ExcelColumnAttribute(), null);
            builder.AppendColumnHeaderRowItem(excelinfo);


        }
        private void PopulateSheetData(List<ActualFlight> list, XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder sheet, ActualWorksheet item)
        {
            NumberFormatInfo format = CultureInfo.CurrentCulture.NumberFormat;
            string currencyFormat = "#,##0.00";

            currencyFormat = currencyFormat.Replace(",", format.CurrencyGroupSeparator).Replace(".", format.CurrencyDecimalSeparator);

            List<ExcelCell> row = new List<ExcelCell>();
            ExcelCell cell = new ExcelCell();

            foreach(var rec in list)
            {
                cell = new ExcelCell();
                cell.CellHeader = item.Data;
                cell.CellValue = rec.DataToString();
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlanName;
                cell.CellValue = rec.PlanName;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.VehicleName;
                cell.CellValue = rec.VehicleName;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ProductName;
                cell.CellValue = rec.ProductName;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.SubtypeName;
                cell.CellValue = rec.SubtypeName;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.UnitName;
                cell.CellValue = rec.UnitName;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.MarketName;
                cell.CellValue = rec.MarketName;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.DemoName;
                cell.CellValue = rec.DemoName;
                row.Add(cell);

                foreach(var it in rec.CustomColumnsValues)
                {
                    cell = new ExcelCell();
                    cell.CellHeader = it.Key;
                    cell.CellValue = it.Value;
                    row.Add(cell);

                }
                if(rec.CustomColumnsValues == null || rec.CustomColumnsValues.Count == 0)
                {
                    cell = new ExcelCell();
                    cell.CellHeader = item.CreativeName;
                    cell.CellValue = rec.CreativeName;
                    row.Add(cell);

                    cell = new ExcelCell();
                    cell.CellHeader = item.FundingSourceName;
                    cell.CellValue = rec.FundingSourceName;
                    row.Add(cell);
                }

                cell = new ExcelCell();
                cell.CellHeader = item.StartDate;
                cell.CellValue = string.Format("{0:MM/dd/yyyy}",rec.StartDate);
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.EndDate;
                cell.CellValue = string.Format("{0:MM/dd/yyyy}",rec.EndDate);
                row.Add(cell);

                //planned
                cell = new ExcelCell();
                cell.CellHeader = item.PlannedGRPs;
                cell.CellValue = rec.Planned.GRPs;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedGross;
                cell.CellValue = rec.Planned.GrossCost;
                if(sheet.SheetColumns.Any(x=>x.Header==item.PlannedGross)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedNet;
                cell.CellValue = rec.Planned.NetCost;
                if (sheet.SheetColumns.Any(x => x.Header == item.PlannedNet)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedProduction;
                cell.CellValue = rec.Planned.Production;
                if (sheet.SheetColumns.Any(x => x.Header == item.PlannedProduction)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedImpressions;
                cell.CellValue = rec.Planned.Impressions;
                if (sheet.SheetColumns.Any(x => x.Header == item.PlannedImpressions)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedClicks;
                cell.CellValue = rec.Planned.Clicks;
                if (sheet.SheetColumns.Any(x => x.Header == item.PlannedClicks)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedLeads;
                cell.CellValue = rec.Planned.Leads;
                if (sheet.SheetColumns.Any(x => x.Header == item.PlannedLeads)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedRichMedia;
                cell.CellValue = rec.Planned.RichMedia;
                if (sheet.SheetColumns.Any(x => x.Header == item.PlannedRichMedia)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedAdServing;
                cell.CellValue = rec.Planned.AdServing;
                if (sheet.SheetColumns.Any(x => x.Header == item.PlannedAdServing)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedTRPs;
                cell.CellValue = rec.Planned.TRPs;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedReach;
                cell.CellValue = rec.Planned.Reach;
                if (sheet.SheetColumns.Any(x => x.Header == item.PlannedReach)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.PlannedFrequency;
                cell.CellValue = rec.Planned.Frequency;
                if (sheet.SheetColumns.Any(x => x.Header == item.PlannedFrequency)) row.Add(cell);

                if (rec.Actualized)
                {
                    cell = new ExcelCell();
                    cell.CellHeader = item.DateActualized;
                    cell.CellValue = string.Format("{0:MM/dd/yyyy}", rec.DateActualized);
                    row.Add(cell);

                    cell = new ExcelCell();
                    cell.CellHeader = item.ActualizedBy;
                    cell.CellValue = rec.ActualizedBy;
                    row.Add(cell);
                }
                else
                {
                    cell.CellHeader = item.DateActualized;
                    cell.CellValue = string.Empty;
                    row.Add(cell);

                    cell = new ExcelCell();
                    cell.CellHeader = item.ActualizedBy;
                    cell.CellValue = string.Empty;
                    row.Add(cell);
                }


                ////actual
                //if (rec.Actual == null)
                //{
                //    sheet.AppendRow(row);
                //    continue;
                //}

                cell = new ExcelCell();
                cell.CellHeader = item.ActualGRPs;
                cell.CellValue = (rec.Actual == null)? 0 : rec.Actual.GRPs;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualGross;                
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.GrossCost;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualGross)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualNet;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.NetCost;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualNet)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualProduction;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.Production;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualProduction)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualImpressions;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.Impressions;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualImpressions)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualClicks;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.Clicks;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualClicks)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualLeads;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.Leads;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualLeads)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualRichMedia;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.RichMedia;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualRichMedia)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualAdServing;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.AdServing;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualAdServing)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualTRPs;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.TRPs;
                row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualReach;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.Reach;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualReach)) row.Add(cell);

                cell = new ExcelCell();
                cell.CellHeader = item.ActualFrequency;
                cell.CellValue = (rec.Actual == null) ? 0 : rec.Actual.Frequency;
                if (sheet.SheetColumns.Any(x => x.Header == item.ActualFrequency)) row.Add(cell);

                sheet.AppendRow(row);
            }

        }
        private void PopulateSheets(IXlsxDocumentBuilder document, ActualExport export)
        {
            foreach (var item in export.Sheets)
            {
                var sheet = new XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder(item.MediaType.Name);
                sheet.ExternalActualsLabel = export.ExternalActualsLabel;
                sheet.ActualRow = true;
                sheet.ColNames = PopulateNameCollection(item);
                // get offset
                var offset = export.Flights.FirstOrDefault(x => x.MediaTypeID == item.MediaType.Id);
                item.SetActualWorksheet(offset.CustomColumnsValues.Count);
                FormatWorkSheet(sheet, item, offset.CustomColumnsValues);
                PopulateSheetData(export.Flights.Where(x=>x.MediaTypeID==item.MediaType.Id).ToList(), sheet, item);
                //HideShowMeasureColumns(sheet, item);
                document.AppendSheet(sheet);
            }

        }
        private Dictionary<string, Tuple<string,bool>> PopulateNameCollection(ActualWorksheet item)
        {
            var names = new Dictionary<string, Tuple<string, bool>>();
            names.Add(item.Data, new Tuple<string, bool>("FlightData", false));
            names.Add(item.ActualGross, new Tuple<string, bool>("ActualGross", true));
            names.Add(item.ActualGRPs, new Tuple<string, bool>(item.ActualGRPs,false));
            names.Add(item.ActualTRPs, new Tuple<string, bool>(item.ActualTRPs,false));
            names.Add(item.ActualNet, new Tuple<string, bool>("ActualNet",true));
            names.Add(item.ActualProduction, new Tuple<string, bool>(item.ActualProduction,true));
            names.Add(item.ActualImpressions, new Tuple<string, bool>(item.ActualImpressions,false));
            names.Add(item.ActualClicks, new Tuple<string, bool>(item.ActualClicks,false));
            names.Add(item.ActualLeads, new Tuple<string, bool>(item.ActualLeads,false));
            names.Add(item.ActualAdServing, new Tuple<string, bool>(item.ActualAdServing,true));
            names.Add(item.ActualRichMedia, new Tuple<string, bool>(item.ActualRichMedia,true));
            names.Add(item.ActualReach, new Tuple<string, bool>(item.ActualReach,false));
            names.Add(item.ActualFrequency, new Tuple<string, bool>(item.ActualFrequency,false));
            names.Add(item.PlannedGross, new Tuple<string, bool>(item.PlannedGross, true));
            names.Add(item.PlannedNet, new Tuple<string, bool>(item.PlannedNet, true));
            names.Add(item.PlannedProduction, new Tuple<string, bool>(item.PlannedProduction, true));
            names.Add(item.PlannedRichMedia, new Tuple<string, bool>(item.PlannedRichMedia, true));
            names.Add(item.PlannedAdServing, new Tuple<string, bool>(item.PlannedAdServing, true));

            return names;
        }
    }
}
