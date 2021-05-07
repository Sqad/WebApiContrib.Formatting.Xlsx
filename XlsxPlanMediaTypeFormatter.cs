using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using OfficeOpenXml.Style;

using SQAD.MTNext.Services.Repositories.Export;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Actuals;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.CostSources;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.DeliverySources;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans;
using WebApiContrib.Formatting.Xlsx.Serialisation.Internal;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted;
using SQAD.XlsxExportImport.Base.Interfaces;
using SQAD.XlsxExportImport.Base.Serialization;
using SQAD.XlsxExportImport.Base.Formatters;
using SQAD.XlsxExportView.Formatted;
using SQAD.XlsxExportView.Unformatted;


namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx
{
    /// <summary>
    /// Class used to send an Excel file to the response.
    /// </summary>
    public class XlsxPlanMediaTypeFormatter : XlsxMediaTypeFormatter
    {
        #region Constructor

        /// <summary>
        /// Create a new ExcelMediaTypeFormatter.
        /// </summary>
        /// <param name="autoFit">True if the formatter should autofit columns after writing the data. (Default
        /// true.)</param>
        /// <param name="autoFilter">True if an autofilter should be applied to the worksheet. (Default false.)</param>
        /// <param name="freezeHeader">True if the header row should be frozen. (Default false.)</param>
        /// <param name="headerHeight">Height of the header row.</param>
        /// <param name="cellHeight">Height of each row of data.</param>
        /// <param name="cellStyle">An action method that modifies the worksheet cell style.</param>
        /// <param name="headerStyle">An action method that modifies the cell style of the first (header) row in the
        /// worksheet.</param>
        public XlsxPlanMediaTypeFormatter(IHttpContextAccessor httpContextAccessor,
                                      IModelMetadataProvider modelMetadataProvider,
                                      bool autoFit = true,
                                      bool autoFilter = false,
                                      bool freezeHeader = false,
                                      double? headerHeight = null,
                                      Action<ExcelStyle> cellStyle = null,
                                      Action<ExcelStyle> headerStyle = null,
                                      SerializerType serializerType = SerializerType.Default,
                                      bool isExportJsonToXls = false,
                                      string fileExtension = null,
                                      string viewLabel = null,
                                      IExportHelpersRepository staticValuesResolver = null) :
            base(httpContextAccessor, modelMetadataProvider, autoFit, autoFilter, freezeHeader,
                  headerHeight, cellStyle, headerStyle, serializerType, isExportJsonToXls, fileExtension)
        {
            //if (string.IsNullOrEmpty(viewLabel))
            //{
            //   viewLabel = httpContextAccessor.HttpContext?.RequestServices.GetService<MTEntitiesContextWithViews>()
            //     ?.DatabaseSettings.Where(x => x.Key.Equals("ViewLabel")).FirstOrDefault()?.Value;
            //}

            // Initialize serializers.
            Serializers = new List<IXlsxSerializer>
                          {
                              new SQADPlanXlsSerializer(staticValuesResolver, modelMetadataProvider, isExportJsonToXls: isExportJsonToXls),
                              new SqadFormattedViewXlsxSerializer(viewLabel),
                              new SqadUnformattedViewXlsxSerializer(viewLabel),
                              new SqadSummaryViewXlsxSerializer(),
                              new SqadActualXlsSerializer(),
                              new SqadCostSourceXlsxSerializer(),
                              new SqadDeliverySourceXlsxSerializer(),
                              new SQADApprovalReportXlsSerializer(),
                              new SqadInternalDatabaseSetupRecordsXlsxSerializer(),
                              new FormattedPlanSerializer()
                          };
        }

        #endregion
    }
}