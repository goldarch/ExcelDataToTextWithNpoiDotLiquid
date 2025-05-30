// GitHub: https://github.com/GoldArch/ExcelDataToTextWithNpoiDotLiquid 
// Version: 1.0.0 (示例版本号)
// Author: GoldArch
// Description: 一个使用 NPOI 读取 Excel 数据并通过 DotLiquid 模板生成文本的小工具。

using DotLiquid;
using System;
using System.Data;
using System.Text;

// It's good practice to put classes into specific namespaces.
namespace ExcelDataToTextTool.TemplateLogic
{
    /// <summary>
    /// Processes DotLiquid templates using data from a DataTable.
    /// </summary>
    public class SqlBuildDotLiquid // Consider renaming if not strictly for SQL, e.g., DataTableTemplateProcessor
    {
        private readonly string templateString;
        private readonly DataTable dataTable;
        private readonly Template parsedTemplate;

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlBuildDotLiquid"/> class.
        /// </summary>
        /// <param name="templateContent">The DotLiquid template string.</param>
        /// <param name="dataSource">The DataTable to use as the data source.</param>
        /// <exception cref="ArgumentNullException">Thrown if templateContent or dataSource is null.</exception>
        /// <exception cref="DotLiquid.Exceptions.SyntaxException">Thrown if the template string has syntax errors.</exception>
        public SqlBuildDotLiquid(string templateContent, DataTable dataSource)
        {
            if (string.IsNullOrWhiteSpace(templateContent))
                throw new ArgumentNullException(nameof(templateContent), "模板内容不能为空。");
            this.dataTable = dataSource ?? throw new ArgumentNullException(nameof(dataSource), "数据源不能为空。");

            this.templateString = templateContent;

            // Pre-parse the template for efficiency if it's used multiple times with different data,
            // or if syntax validation at construction is desired.
            try
            {
                this.parsedTemplate = Template.Parse(this.templateString);
            }
            catch (DotLiquid.Exceptions.SyntaxException ex)
            {
                // Optionally rethrow with more context or handle as needed
                throw new DotLiquid.Exceptions.SyntaxException($"解析模板时发生语法错误: {ex.Message}");
            }
        }

        /// <summary>
        /// Generates text by applying the template to each row of the DataTable.
        /// Each row is available in the template контекст as 'row'.
        /// </summary>
        /// <returns>The generated text, with results from each row appended.</returns>
        public string GenerateTextFromDataRowTemplate()
        {
            if (this.dataTable.Rows.Count == 0)
            {
                return string.Empty; // Or a message indicating no data
            }

            StringBuilder sb = new StringBuilder();

            foreach (DataRow dataRow in this.dataTable.Rows)
            {
                // The Hash object creates the root scope for the template rendering.
                // We're making the DataRowDrop available under the name 'row'.
                //var renderParameters = new RenderParameters(System.Globalization.CultureInfo.InvariantCulture)
                var renderParameters = new RenderParameters()
                {
                    LocalVariables = Hash.FromAnonymousObject(new { row = new DataRowDrop(dataRow) })
                    // Filters can be registered globally or per render call if needed
                    // Example: Filters = new[] { typeof(MyCustomFilters) }
                };

                string renderedRow = this.parsedTemplate.Render(renderParameters);
                sb.AppendLine(renderedRow);
            }
            return sb.ToString();
        }

        /// <summary>
        /// A DotLiquid Drop to expose DataRow fields to the template.
        /// Access columns using {{ row.ColumnName }}.
        /// </summary>
        internal class DataRowDrop : Drop
        {
            private readonly DataRow _dataRow;

            public DataRowDrop(DataRow dr)
            {
                this._dataRow = dr ?? throw new ArgumentNullException(nameof(dr));
            }

            /// <summary>
            /// Called by DotLiquid when a property (column name) is accessed on this drop.
            /// </summary>
            /// <param name="methodOrPropertyName">The name of the column to access.</param>
            /// <returns>The value of the column, or an empty string if the column doesn't exist or its value is null/whitespace.</returns>
            public override object BeforeMethod(string methodOrPropertyName)
            {
                if (this._dataRow.Table.Columns.Contains(methodOrPropertyName))
                {
                    object cellValue = this._dataRow[methodOrPropertyName];
                    // Handle DBNull explicitly, return empty string for null or whitespace strings.
                    if (cellValue == DBNull.Value || cellValue == null)
                    {
                        return string.Empty;
                    }
                    // If you have a DbValidate.IsNullOrWhiteSpace equivalent:
                    // if (YourNamespace.DbValidate.IsNullOrWhiteSpace(cellValue)) return string.Empty;
                    // Otherwise, a simple check for string:
                    if (cellValue is string s && string.IsNullOrWhiteSpace(s))
                    {
                        return string.Empty;
                    }
                    return cellValue;
                }
                // If column does not exist, DotLiquid typically returns null, which renders as empty string.
                // Returning empty string explicitly is also fine.
                return string.Empty;
            }
        }
    }
}
