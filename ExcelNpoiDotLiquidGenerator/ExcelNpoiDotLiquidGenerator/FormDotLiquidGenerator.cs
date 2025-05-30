// FormDotLiquidGenerator.cs
// GitHub: https://github.com/GoldArch/ExcelDataToTextWithNpoiDotLiquid 
// Version: 1.0.0 (示例版本号)
// Author: GoldArch
// Description: 一个使用 NPOI 读取 Excel 数据并通过 DotLiquid 模板生成文本的小工具。

using System;
using System.Data;
using System.Windows.Forms;
using System.IO; // Required for NPOI FileStream and Path operations

// Add NPOI using statements
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel; // For .xls files
using NPOI.XSSF.UserModel; // For .xlsx files

namespace ExcelDataToTextTool
{
    public partial class FormDotLiquidGenerator : Form
    {
        #region Private Fields
        private DataTable importedDataTable;
        #endregion

        #region UI Control Fields
        private Button btnLoadExcel;
        private Button btnGenerateRowByRow;
        private Button btnCopyOutput;
        private TextBox txtTemplate;
        private TextBox txtOutput;
        private DataGridView dgvImportedData;
        private Label lblTemplate;
        private Label lblOutput;
        private Label lblImportedData;
        #endregion

        #region Constructor
        public FormDotLiquidGenerator()
        {
            InitializeComponent();
        }
        #endregion

        #region UI Initialization
        private void InitializeComponent()
        {
            this.btnLoadExcel = new System.Windows.Forms.Button();
            this.btnGenerateRowByRow = new System.Windows.Forms.Button();
            this.btnCopyOutput = new System.Windows.Forms.Button();
            this.txtTemplate = new System.Windows.Forms.TextBox();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.dgvImportedData = new System.Windows.Forms.DataGridView();
            this.lblTemplate = new System.Windows.Forms.Label();
            this.lblOutput = new System.Windows.Forms.Label();
            this.lblImportedData = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvImportedData)).BeginInit();
            this.SuspendLayout();
            // 
            // btnLoadExcel
            // 
            this.btnLoadExcel.Location = new System.Drawing.Point(12, 12);
            this.btnLoadExcel.Name = "btnLoadExcel";
            this.btnLoadExcel.Size = new System.Drawing.Size(130, 30);
            this.btnLoadExcel.TabIndex = 0;
            this.btnLoadExcel.Text = "导入 Excel 文件";
            this.btnLoadExcel.UseVisualStyleBackColor = true;
            this.btnLoadExcel.Click += new System.EventHandler(this.btnLoadExcel_Click);
            // 
            // btnGenerateRowByRow
            // 
            this.btnGenerateRowByRow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnGenerateRowByRow.Location = new System.Drawing.Point(12, 440);
            this.btnGenerateRowByRow.Name = "btnGenerateRowByRow";
            this.btnGenerateRowByRow.Size = new System.Drawing.Size(150, 30);
            this.btnGenerateRowByRow.TabIndex = 3;
            this.btnGenerateRowByRow.Text = "生成文本 (逐行)";
            this.btnGenerateRowByRow.UseVisualStyleBackColor = true;
            this.btnGenerateRowByRow.Click += new System.EventHandler(this.btnGenerateRowByRow_Click);
            // 
            // btnCopyOutput
            // 
            this.btnCopyOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCopyOutput.Location = new System.Drawing.Point(232, 440);
            this.btnCopyOutput.Name = "btnCopyOutput";
            this.btnCopyOutput.Size = new System.Drawing.Size(150, 30);
            this.btnCopyOutput.TabIndex = 4;
            this.btnCopyOutput.Text = "复制结果";
            this.btnCopyOutput.UseVisualStyleBackColor = true;
            this.btnCopyOutput.Click += new System.EventHandler(this.btnCopyOutput_Click);
            // 
            // txtTemplate
            // 
            this.txtTemplate.AcceptsReturn = true;
            this.txtTemplate.AcceptsTab = true;
            this.txtTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtTemplate.Location = new System.Drawing.Point(12, 280);
            this.txtTemplate.Multiline = true;
            this.txtTemplate.Name = "txtTemplate";
            this.txtTemplate.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtTemplate.Size = new System.Drawing.Size(370, 150);
            this.txtTemplate.TabIndex = 2;
            // 
            // txtOutput
            // 
            this.txtOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOutput.Location = new System.Drawing.Point(400, 280);
            this.txtOutput.Multiline = true;
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ReadOnly = true;
            this.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtOutput.Size = new System.Drawing.Size(370, 190);
            this.txtOutput.TabIndex = 5;
            // 
            // dgvImportedData
            // 
            this.dgvImportedData.AllowUserToAddRows = false;
            this.dgvImportedData.AllowUserToDeleteRows = false;
            this.dgvImportedData.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvImportedData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvImportedData.Location = new System.Drawing.Point(12, 70);
            this.dgvImportedData.Name = "dgvImportedData";
            this.dgvImportedData.ReadOnly = true;
            this.dgvImportedData.RowTemplate.Height = 23;
            this.dgvImportedData.Size = new System.Drawing.Size(760, 180);
            this.dgvImportedData.TabIndex = 1;
            // 
            // lblTemplate
            // 
            this.lblTemplate.AutoSize = true;
            this.lblTemplate.Location = new System.Drawing.Point(12, 260);
            this.lblTemplate.Name = "lblTemplate";
            this.lblTemplate.Size = new System.Drawing.Size(95, 12);
            this.lblTemplate.TabIndex = 7;
            this.lblTemplate.Text = "DotLiquid 模板:";
            // 
            // lblOutput
            // 
            this.lblOutput.AutoSize = true;
            this.lblOutput.Location = new System.Drawing.Point(397, 260);
            this.lblOutput.Name = "lblOutput";
            this.lblOutput.Size = new System.Drawing.Size(59, 12);
            this.lblOutput.TabIndex = 8;
            this.lblOutput.Text = "生成结果:";
            // 
            // lblImportedData
            // 
            this.lblImportedData.AutoSize = true;
            this.lblImportedData.Location = new System.Drawing.Point(12, 50);
            this.lblImportedData.Name = "lblImportedData";
            this.lblImportedData.Size = new System.Drawing.Size(71, 12);
            this.lblImportedData.TabIndex = 6;
            this.lblImportedData.Text = "导入的数据:";
            // 
            // FormDotLiquidGenerator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 481);
            this.Controls.Add(this.lblOutput);
            this.Controls.Add(this.lblTemplate);
            this.Controls.Add(this.lblImportedData);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.btnCopyOutput);
            this.Controls.Add(this.btnGenerateRowByRow);
            this.Controls.Add(this.txtTemplate);
            this.Controls.Add(this.dgvImportedData);
            this.Controls.Add(this.btnLoadExcel);
            this.MinimumSize = new System.Drawing.Size(600, 400);
            this.Name = "FormDotLiquidGenerator";
            this.Text = "Excel 数据转文本工具 (NPOI + DotLiquid) - Developed by GoldArch";
            ((System.ComponentModel.ISupportInitialize)(this.dgvImportedData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        #region Event Handlers
        private void btnLoadExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = @"Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*", // Prioritize .xlsx
                Title = "选择 Excel 文件"
            })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        string filePath = dlg.FileName;
                        this.importedDataTable = ExcelToDataTable(filePath, true); // Assuming first row is header

                        if (this.importedDataTable != null)
                        {
                            this.dgvImportedData.DataSource = this.importedDataTable;
                            MessageBox.Show($"成功从 '{Path.GetFileName(filePath)}' 加载 {this.importedDataTable.Rows.Count} 行数据。",
                                            "数据加载成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            this.dgvImportedData.DataSource = null;
                            // ExcelToDataTable should show specific error messages.
                        }
                    }
                    catch (Exception ex)
                    {
                        this.importedDataTable = null;
                        this.dgvImportedData.DataSource = null;
                        MessageBox.Show($"加载 Excel 文件时发生意外错误: {ex.Message}\n\n{ex.StackTrace}",
                                        "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnGenerateRowByRow_Click(object sender, EventArgs e)
        {
            if (this.importedDataTable == null || this.importedDataTable.Rows.Count == 0)
            {
                MessageBox.Show("请先加载包含数据的 Excel 文件。", "无数据", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(this.txtTemplate.Text))
            {
                MessageBox.Show("请输入 DotLiquid 模板。", "无模板", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.txtTemplate.Focus();
                return;
            }

            try
            {
                // Ensure SqlBuildDotLiquid class is accessible (e.g., same namespace or via using directive)
                var templateProcessor = new ExcelDataToTextTool.TemplateLogic.SqlBuildDotLiquid(this.txtTemplate.Text, this.importedDataTable);
                this.txtOutput.Text = templateProcessor.GenerateTextFromDataRowTemplate();
                MessageBox.Show("文本生成完成。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (DotLiquid.Exceptions.SyntaxException syntaxEx)
            {
                this.txtOutput.Text = $"模板语法错误: {syntaxEx.Message}\n\n详细信息:\n{syntaxEx.StackTrace}";
                MessageBox.Show($"模板语法错误: {syntaxEx.Message}", "模板错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                this.txtOutput.Text = $"生成输出时出错: {ex.Message}\n\n详细信息:\n{ex.StackTrace}";
                MessageBox.Show($"生成文本时出错: {ex.Message}", "生成错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCopyOutput_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtOutput.Text))
            {
                try
                {
                    Clipboard.SetText(this.txtOutput.Text);
                    MessageBox.Show("结果已复制到剪贴板。", "复制成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (System.Runtime.InteropServices.ExternalException ex)
                {
                    MessageBox.Show($"无法访问剪贴板: {ex.Message}", "复制失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("没有可复制的内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion

        #region Helper Methods
        /// <summary>
        /// Reads data from an Excel file into a DataTable using NPOI.
        /// Assumes the first sheet is used.
        /// </summary>
        /// <param name="filePath">The full path to the Excel file.</param>
        /// <param name="hasHeaderRow">Indicates if the first row in Excel is a header row.</param>
        /// <returns>A DataTable containing the Excel data, or null if an error occurs.</returns>
        public static DataTable ExcelToDataTable(string filePath, bool hasHeaderRow)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) // Allow read sharing
                {
                    string fileExt = Path.GetExtension(filePath)?.ToLowerInvariant(); // Use ToLowerInvariant for consistency
                    if (fileExt == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    else if (fileExt == ".xls")
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    else
                    {
                        MessageBox.Show("不支持的文件格式。请选择 .xls 或 .xlsx 文件。", "格式错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }
                }
            }
            catch (IOException ioEx) // More specific exception for file access issues
            {
                MessageBox.Show($"打开 Excel 文件时发生 IO 错误 (文件可能被占用或路径无效): {ioEx.Message}", "文件读取错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            catch (Exception ex) // General exception for other NPOI loading issues
            {
                MessageBox.Show($"打开 Excel 文件时出错: {ex.Message}", "文件读取错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            if (workbook.NumberOfSheets == 0)
            {
                MessageBox.Show("Excel 文件不包含任何工作表。", "文件内容错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            ISheet sheet = workbook.GetSheetAt(0); // Get the first sheet
            if (sheet == null || sheet.PhysicalNumberOfRows == 0) // Check if sheet is null or empty
            {
                MessageBox.Show("Excel 文件中的第一个工作表为空或无效。", "文件内容错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            // Determine the starting row for data and header
            int firstDataRowIndex = sheet.FirstRowNum;
            IRow headerSourceRow = sheet.GetRow(firstDataRowIndex);

            if (headerSourceRow == null) // If the very first row is null, unlikely but possible for sparse files
            {
                MessageBox.Show("无法读取Excel文件的行数据。", "文件内容错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }


            // Create columns in DataTable
            if (hasHeaderRow)
            {
                if (sheet.GetRow(firstDataRowIndex) == null)
                { // Ensure header row exists
                    MessageBox.Show("指定的标题行不存在或为空。", "文件内容错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }
                for (int i = headerSourceRow.FirstCellNum; i < headerSourceRow.LastCellNum; i++)
                {
                    ICell cell = headerSourceRow.GetCell(i);
                    string columnName = (cell == null || string.IsNullOrWhiteSpace(cell.ToString()))
                                        ? $"列{i + 1}" // Use Chinese for default column name
                                        : cell.ToString().Trim();
                    // Handle duplicate column names by appending a unique suffix
                    int duplicateCount = 1;
                    string originalColumnName = columnName;
                    while (dt.Columns.Contains(columnName))
                    {
                        columnName = $"{originalColumnName}_{duplicateCount++}";
                    }
                    dt.Columns.Add(new DataColumn(columnName));
                }
                firstDataRowIndex++; // Data starts from the next row
            }
            else
            {
                // Auto-generate column names if no header row
                if (headerSourceRow.LastCellNum <= 0 && sheet.PhysicalNumberOfRows > 0)
                { // If first row has no cells, try to find a row with cells to determine count
                    for (int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
                    {
                        IRow tempRow = sheet.GetRow(r);
                        if (tempRow != null && tempRow.LastCellNum > 0)
                        {
                            headerSourceRow = tempRow;
                            break;
                        }
                    }
                }
                if (headerSourceRow.LastCellNum <= 0)
                {
                    MessageBox.Show("无法确定Excel文件的列数。", "文件内容错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }

                for (int i = 0; i < headerSourceRow.LastCellNum; i++)
                {
                    dt.Columns.Add(new DataColumn($"列{i + 1}")); // Use Chinese for default column name
                }
            }

            // Populate data rows
            for (int i = firstDataRowIndex; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue; // Skip empty rows

                DataRow dataRow = dt.NewRow();
                bool allCellsEmptyInRow = true;

                for (int j = 0; j < dt.Columns.Count; j++) // Iterate based on DataTable columns
                {
                    // Adjust cell index if Excel row starts from a non-zero FirstCellNum
                    ICell cell = row.GetCell(j + row.FirstCellNum);
                    object cellValue = null;

                    if (cell != null)
                    {
                        switch (cell.CellType)
                        {
                            case CellType.String:
                                cellValue = cell.StringCellValue;
                                break;
                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(cell))
                                    cellValue = cell.DateCellValue;
                                else
                                    cellValue = cell.NumericCellValue;
                                break;
                            case CellType.Boolean:
                                cellValue = cell.BooleanCellValue;
                                break;
                            case CellType.Formula:
                                try
                                {
                                    IFormulaEvaluator evaluator = WorkbookFactory.CreateFormulaEvaluator(workbook);
                                    // For HSSF (.xls), sometimes EvaluateInCell is needed before Evaluate
                                    if (workbook is HSSFWorkbook) evaluator.EvaluateInCell(cell);
                                    CellValue evaluatedCv = evaluator.Evaluate(cell);
                                    switch (evaluatedCv.CellType)
                                    {
                                        case CellType.String: cellValue = evaluatedCv.StringValue; break;
                                        case CellType.Numeric:
                                            // For date formatted formula cells, check original cell's format
                                            if (DateUtil.IsCellDateFormatted(cell)) cellValue = cell.DateCellValue;
                                            else cellValue = evaluatedCv.NumberValue;
                                            break;
                                        case CellType.Boolean: cellValue = evaluatedCv.BooleanValue; break;
                                        default: cellValue = cell.ToString(); break; // Fallback
                                    }
                                }
                                catch
                                {
                                    try { cellValue = cell.CellFormula; } // Fallback to formula string
                                    catch { cellValue = cell.ToString(); } // Final fallback
                                }
                                break;
                            case CellType.Blank: // Treat blank as DBNull
                            case CellType.Unknown:
                            case CellType.Error:
                                cellValue = DBNull.Value;
                                break;
                            default:
                                cellValue = cell.ToString(); // Fallback for other types
                                break;
                        }

                        if (cellValue != null && cellValue != DBNull.Value && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        {
                            allCellsEmptyInRow = false;
                        }
                        dataRow[j] = cellValue ?? DBNull.Value; // Ensure DBNull if cellValue is null
                    }
                    else
                    {
                        dataRow[j] = DBNull.Value;
                    }
                }

                if (!allCellsEmptyInRow)
                {
                    dt.Rows.Add(dataRow);
                }
            }
            return dt;
        }
        #endregion
    }
}
