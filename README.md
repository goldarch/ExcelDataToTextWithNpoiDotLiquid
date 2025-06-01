# ExcelDataToTextWithNpoiDotLiquid
.net4 
适用于逐行数据嵌套到模板中渲染的场景，比如用一行数据嵌套到sql语句模板中，生成sql语言；又比如，用户信息嵌套入一个标准格式的模板中，生成格式化的邮件内容等
特别是企业中，有导入数据的场景，可以通过sql模板，可以高度自定义的生成导入模板，如果是临时性的，可以复制生成的脚本到数据库管理工具中去执行。如果是固化的经常需要导入的内容，可以固化模板，让用户的引入excel导入特定数据（需要增加数据检验和与数据库的连接） 
## **软件截图**
![image](https://github.com/goldarch/ExcelDataToTextWithNpoiDotLiquid/blob/main/%E5%9B%BE%E7%89%87%E8%B5%84%E6%BA%90/%E6%B5%8B%E8%AF%95%E7%9A%84%E6%95%88%E6%9E%9C.png)  

作者: GoldArch  
GitHub 仓库: https://github.com/goldarch/ExcelDataToTextWithNpoiDotLiquid  
版本: 1.0.0 (示例)  
一款基于 .NET Windows Forms 的实用小工具，它使用 NPOI 库读取 Excel 文件（.xls 和 .xlsx 格式），并结合 DotLiquid 模板引擎，帮助用户根据 Excel 表格中的数据逐行生成自定义的文本内容。

## **主要功能**

* **导入 Excel 数据**: 支持导入 .xls 和 .xlsx 格式的 Excel 文件，并默认读取第一个工作表的数据。  
* **数据显示**: 在界面上清晰展示导入的 Excel 数据，方便用户预览。  
* **DotLiquid 模板驱动**: 用户可以编写 DotLiquid 模板，灵活定义输出文本的格式。  
* **逐行文本生成**: 工具会遍历 Excel 表格的每一行数据，并将模板应用于该行，生成相应的文本。  
* **结果预览与复制**: 生成的文本会显示在输出框中，并提供一键复制到剪贴板的功能。
## **系统要求**

* **操作系统**: Windows  
* **.NET Framework**: .NET Framework 4.0 或更高版本

## **依赖库 (NuGet Packages)**

本项目主要依赖以下 NuGet 包：

* **NPOI**: 用于读取 Excel 文件。  
* **DotLiquid**: 用于处理模板生成。

在 Visual Studio 中打开项目时，这些依赖项通常会自动还原。

## **如何使用**

1. **获取软件**:  
   * 您可以从本仓库的 [Debug页面]([https://github.com/goldarch/ExcelDataToTextWithNpoiDotLiquid/releases](https://github.com/goldarch/ExcelDataToTextWithNpoiDotLiquid/tree/main/ExcelNpoiDotLiquidGenerator/ExcelNpoiDotLiquidGenerator/bin/Debug)) 下载最新的已编译版本 
   * 或者，您可以克隆本仓库并自行编译（参见下面的“从源码编译”部分）。  
2. **运行程序**: 双击运行 ExcelDataToTextTool.exe (或您编译后的可执行文件名)。  
3. **导入 Excel 文件**:  
   * 点击界面左上角的 “导入 Excel 文件” 按钮。  
   * 在弹出的对话框中选择您的 Excel 文件。  
   * 导入成功后，数据将显示在主界面的表格中（默认读取第一个 Sheet，并将第一行作为表头）。  
4. **编写 DotLiquid 模板**:  
   * 在左下角的 “DotLiquid 模板:” 文本框中输入您的模板。  
   * 模板中可以使用 {{ row.列名 }} 的形式来引用当前行对应列的数据。**请确保 列名 与您 Excel 文件中的表头名称完全一致（包括大小写，如果您的表头是中文，则使用中文列名）。**  
5. **生成文本**:  
   * 点击 “生成文本 (逐行)” 按钮。  
6. **查看并复制结果**:  
   * 生成的文本将显示在右侧的 “生成结果:” 文本框中。  
   * 点击 “复制结果” 按钮，可以将生成的全部文本复制到剪贴板。

## **DotLiquid 模板示例**

假设您的 Excel 文件有以下列：ID, 产品名称, 数量, 单价

**示例模板 1 (生成逗号分隔值):**

{{ row.ID }},{{ row.产品名称 }},{{ row.数量 }},{{ row.单价 }}

**示例模板** 2 (生成 **SQL 插入语句的一部分 \- 请注意SQL注入风险，仅作示例):**

INSERT INTO Products (ID, Name, Quantity, Price) VALUES ('{{ row.ID }}', '{{ row.产品名称 }}', {{ row.数量 }}, {{ row.单价 }});

**示例模板 3 (生成 Markdown 列表项):**

\- 产品ID: {{ row.ID }}  
  \- 名称: {{ row.产品名称 }}  
  \- 数量: {{ row.数量 }}  
  \- 单价: {{ row.单价 }}

**重要提示:**

* 模板中的 row. 后面的名称必须与 Excel 表头完全匹配。  
* 如果 Excel 表头包含空格或特殊字符，DotLiquid 可能无法直接通过点表示法访问，建议在导入前规范化 Excel 表头。  
* DataRowDrop 类在处理时，如果单元格值为 DBNull、null 或仅包含空白字符的字符串，则在模板中对应的值会是空字符串 ""。

## **从源码编译 (可选)**

如果您希望自行编译项目：

1. **克隆仓库**:  
   git clone \[https://github.com/goldarch/ExcelDataToTextWithNpoiDotLiquid.git\](https://github.com/goldarch/ExcelDataToTextWithNpoiDotLiquid.git)

2. **打开解决方案**: 使用 Visual Studio (推荐 2017 或更高版本) 打开 ExcelDataToTextWithNpoiDotLiquid.sln (或您的解决方案文件名)。  
3. **还原** NuGet **包**: 在 Visual Studio 中，右键点击解决方案 \-\> “还原 NuGet 程序包”。  
4. **编译项目**: 生成解决方案 (通常是按 F6 或通过菜单 “生成” \-\> “生成解决方案”)。  
5. 生成的可执行文件通常位于项目的 bin\\Debug 或 bin\\Release 目录下。

## **许可证**
This project is licensed under the MIT License
## **贡献代码 (可选)**
欢迎对本项目进行贡献！如果您有任何改进建议或发现了 Bug，请随时提交 Pull Request 或创建 Issue。
## **致谢**

* 本项目使用了优秀的开源库 [NPOI](https://github.com/nissl-lab/npoi) 来处理 Excel 文件。  
* 本项目使用了强大的模板引擎 [DotLiquid](https://github.com/dotliquid/dotliquid) 来实现文本生成。
