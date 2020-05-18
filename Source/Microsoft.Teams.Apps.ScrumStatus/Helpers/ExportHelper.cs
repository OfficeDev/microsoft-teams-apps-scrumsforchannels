// <copyright file="ExportHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ScrumStatus.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Reflection;
    using ClosedXML.Excel;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Helper class for methods to Export to Excel
    /// </summary>
    public class ExportHelper
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<ExportHelper> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExportHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public ExportHelper(ILogger<ExportHelper> logger)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Method to create excel workbook and return the output as stream
        /// </summary>
        /// <param name="exportDt">Excel to export</param>
        /// <returns>stream of file generated</returns>
        public MemoryStream ExportToExcel(DataTable exportDt)
        {
            try
            {
                using (XLWorkbook workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add(exportDt);
                    using (MemoryStream stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        return stream;
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error creating the workbook while exporting the data at ExportToExcel: {ex}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Generic Method to convert list to DataTable.
        /// </summary>
        /// <typeparam name="T">Any generic type.</typeparam>
        /// <param name="scrumToExport">List of scrum to export.</param>
        /// <param name="dataTableName">Name of the data table.</param>
        /// <returns>Return a data table with scrum data.</returns>
        public DataTable ConvertToDataTable<T>(List<T> scrumToExport, string dataTableName)
        {
            try
            {
                DataTable dataTable = new DataTable(dataTableName);

                // Get all the properties.
                PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                foreach (PropertyInfo prop in props)
                {
                    // Defining type of data column gives proper data table.
                    var type = prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType;

                    // Setting column names as Property names.
                    dataTable.Columns.Add(prop.Name, type);
                }

                if (scrumToExport == null)
                {
                    return dataTable;
                }

                foreach (T scrum in scrumToExport)
                {
                    var values = new object[props.Length];
                    for (int i = 0; i < props.Length; i++)
                    {
                        // inserting property values to data table rows.
                        values[i] = props[i].GetValue(scrum, null);
                    }

                    dataTable.Rows.Add(values);
                }

                return dataTable;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while creating data table at ConvertToDataTable: {ex}", SeverityLevel.Error);
                throw;
            }
        }
    }
}
