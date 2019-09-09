using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using SQAD.MTNext.Business.Models.Core.CostSource;
using SQAD.MTNext.Business.Models.Core.Demo;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.CostSources
{
    public class SqadCostSourceXlsxSerializer : IXlsxSerialiser
    {
        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(CostSourceExportDataModel);
        }

        public void Serialise(Type itemType,
                              object value,
                              IXlsxDocumentBuilder document,
                              string sheetName,
                              string columnPrefix,
                              SqadXlsxPlanSheetBuilder sheetbuilderOverride)
        {
            if (!(value is CostSourceExportDataModel exportData))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var dataTable = CreateDataTable(exportData);
            var sheetBuilder = new SqadCostSourceDataSheetBuilder("Data", dataTable, exportData.CostPeriods.Count);
            document.AppendSheet(sheetBuilder);
        }

        private static DataTable CreateDataTable(CostSourceExportDataModel exportData)
        {
            var table = new DataTable();

            table.Columns.Add("market", typeof(string)).Caption = "MARKET";
            table.Columns.Add("subtype", typeof(string)).Caption = "SUBTYPE";
            if (exportData.NeedIncludeDemos)
            {
                table.Columns.Add("demo", typeof(string)).Caption = "DEMO";
            }
            table.Columns.Add("unit", typeof(string)).Caption = "UNIT";

            foreach (var period in exportData.CostPeriods)
            {
                table.Columns.Add($"period_{period.ID}", typeof(decimal)).Caption = period.Caption ?? period.PendingCaption;
            }

            var valuesLookup = exportData.CostValues
                                         .ToDictionary(x => $"{x.MarketID}-{x.AudienceID}-{x.DemoID}-{x.UnitID}-{x.PeriodID}",
                                                       x => x.Cost);

            foreach (var market in exportData.Markets)
            {
                foreach (var subtype in exportData.Subtypes)
                {
                    var demos = exportData.Demos;
                    if (!exportData.NeedIncludeDemos)
                    {
                        demos = new List<DemoModel>
                                {
                                    new DemoModel
                                    {
                                        ID = 0,
                                        Name = ""
                                    }
                                };
                    }

                    foreach (var demo in demos)
                    {
                        foreach (var unit in exportData.Units)
                        {
                            var dataRow = table.NewRow();

                            dataRow["market"] = market.Name;
                            dataRow["subtype"] = subtype.Name;

                            if (exportData.NeedIncludeDemos)
                            {
                                dataRow["demo"] = demo.Name;
                            }

                            dataRow["unit"] = unit.Name;

                            foreach (var period in exportData.CostPeriods)
                            {
                                var key = $"{market.ID}-{subtype.ID}-{demo.ID}-{unit.ID}-{period.ID}";
                                if (!valuesLookup.ContainsKey(key))
                                {
                                    continue;
                                }

                                var value = valuesLookup[key];
                                dataRow[$"period_{period.ID}"] = value;
                            }

                            table.Rows.Add(dataRow);
                        }
                    }
                }
            }

            return table;
        }
    }
}
