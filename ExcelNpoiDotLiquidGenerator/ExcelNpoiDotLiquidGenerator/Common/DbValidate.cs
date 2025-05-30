using System;
using System.Numerics;

namespace ExcelDataToTextTool.Common
{
    public partial class DbValidate
    {
        public static bool IsNullOrWhiteSpaceOrZeroOrFalse(object fieldValue)
        {
            //https://docs.microsoft.com/zh-cn/dotnet/csharp/language-reference/builtin-types/nullable-value-types
            //判断可空类型
            //bool IsNullable(Type type) => Nullable.GetUnderlyingType(type) != null;

            //不论何种类型类型，要是空的，都可以返回。
            if (fieldValue == null)
            {
                return true;
            }

            if (IsNumeric(fieldValue))
            {
                //Float - 7 digits(32 bit)
                //Double - 15 - 16 digits(64 bit)
                //Decimal - 28 - 29 significant digits(128 bit)
                //
                //dx通过tostring，减少了各种类型的转换
                //此处已经包含布尔值
                return decimal.Parse(fieldValue.ToString()) == 0m;
            }


            //DataColumn.DataType Property
            //https://docs.microsoft.com/zh-cn/dotnet/api/system.data.datacolumn.datatype?view=netcore-3.1


            //【重要】不为空的，只判断以下几项，其它都返回非空
            bool retu = false;
            switch (fieldValue.GetType().Name.ToLower())
            {
                case "dbnull":
                    //str = "NULL";
                    retu = true;
                    break;
                case "string":
                    //str = "'" + ((string)someValue).Replace("'", "''") + "'";
                    retu = string.IsNullOrWhiteSpace(fieldValue.ToString());
                    break;
                case "guid": //guid的可空类型好象也是引流到这里，没有真正测试
                    retu = (Guid)fieldValue == Guid.Empty;
                    break;
                case "guid?": //
                    retu = (Guid?)fieldValue == Guid.Empty;
                    break;
                //default:
                //throw new ArgumentOutOfRangeException(@"数据","数据类型的判断超出范围，请与开发人员联系");
            }

            return retu;
        }

        public static bool IsNullOrWhiteSpaceOrZeroOrFalse(object fieldValue, out string pReport)
        {
            //https://docs.microsoft.com/zh-cn/dotnet/csharp/language-reference/builtin-types/nullable-value-types
            //判断可空类型
            //bool IsNullable(Type type) => Nullable.GetUnderlyingType(type) != null;

            pReport = "";

            //不论何种类型类型，要是空的，都可以返回。
            if (fieldValue == null)
            {
                pReport = "值为Null";
                return true;
            }

            if (IsNumeric(fieldValue))
            {
                //Float - 7 digits(32 bit)
                //Double - 15 - 16 digits(64 bit)
                //Decimal - 28 - 29 significant digits(128 bit)
                //
                //dx通过tostring，减少了各种类型的转换
                //此处已经包含布尔值

                pReport = "值为0";

                return decimal.Parse(fieldValue.ToString()) == 0m;
            }


            //DataColumn.DataType Property
            //https://docs.microsoft.com/zh-cn/dotnet/api/system.data.datacolumn.datatype?view=netcore-3.1


            //【重要】不为空的，只判断以下几项，其它都返回非空
            bool retu = false;
            switch (fieldValue.GetType().Name.ToLower())
            {
                case "dbnull":
                    //str = "NULL";
                    pReport = "值为DBNull";
                    retu = true;
                    break;
                case "string":
                    //str = "'" + ((string)someValue).Replace("'", "''") + "'";
                    pReport = "字符串为空";
                    retu = string.IsNullOrWhiteSpace(fieldValue.ToString());
                    break;
                case "guid": //guid的可空类型好象也是引流到这里，没有真正测试
                    pReport = "Guid为Empty";
                    retu = (Guid)fieldValue == Guid.Empty;
                    break;
                case "guid?": //
                    pReport = "Guid?为Empty";
                    retu = (Guid?)fieldValue == Guid.Empty;
                    break;
                //default:
                //throw new ArgumentOutOfRangeException(@"数据","数据类型的判断超出范围，请与开发人员联系");
            }

            return retu;
        }


        //兼容历史代码
        public static bool IsNullOrWhiteSpace(object fieldValue)
        {
            //dx2025.01.05 不能使用IsNullOrWhiteSpaceOrZeroOrFalse,特别是引入excel数据时，不能把0作为空的判断返回！
            //==========================================================
            //比如这里：“空”和“0”是两个完全不同的概念！就不能用IsNullOrWhiteSpaceOrZeroOrFalse把两者当一个东西看！
            //【特定导入规则】：只修改有数据的cell，没有数据的不进行判断
            //if (!DbValidate.IsNullOrWhiteSpace(dataDt.Rows[i]["失业金"]))
            //又如：d:\twsoft\测试代码\员工信息导入\SqlBuild_DotLiquid.cs，这里引入的时候空值和0也是完全不同！
            //return DbValidate.IsNullOrWhiteSpace(_dataRow[method]) ? "" : _dataRow[method];
            //==========================================================

            //return IsNullOrWhiteSpaceOrZeroOrFalse(fieldValue);

            //https://docs.microsoft.com/zh-cn/dotnet/csharp/language-reference/builtin-types/nullable-value-types
            //判断可空类型
            //bool IsNullable(Type type) => Nullable.GetUnderlyingType(type) != null;

            //不论何种类型类型，要是空的，都可以返回。
            if (fieldValue == null)
            {
                return true;
            }

            //【重要】不为空的，只判断以下几项，其它都返回非空
            bool returnVal = false;
            switch (fieldValue.GetType().Name.ToLower())
            {
                case "dbnull":
                    //str = "NULL";
                    returnVal = true;
                    break;
                case "string":
                    //str = "'" + ((string)someValue).Replace("'", "''") + "'";
                    returnVal = string.IsNullOrWhiteSpace(fieldValue.ToString());
                    break;
                case "guid": //guid的可空类型好象也是引流到这里，没有真正测试
                    returnVal = (Guid)fieldValue == Guid.Empty;
                    break;
                case "guid?": //
                    returnVal = (Guid?)fieldValue == Guid.Empty;
                    break;
                //default:
                //throw new ArgumentOutOfRangeException(@"数据","数据类型的判断超出范围，请与开发人员联系");
            }

            //几个特定类型判断是否空后，其它的类型不进行判断
            return returnVal;
        }


        //dx20200627 学习PGK的方式，这样调用的时候不用！号，直观，输入更方便
        public static bool IsNotNullOrNotWhiteSpace(object fieldValue)
        {
            return !IsNullOrWhiteSpace(fieldValue);
        }

        //https://docs.microsoft.com/en-us/dotnet/api/system.valuetype?view=netframework-4.0

        //Value Type and Reference Type
        //https://www.tutorialsteacher.com/csharp/csharp-value-type-and-reference-type
        //这是微软方法，但是，ValueType的求值未研究
        //public static bool IsNumeric(ValueType value)
        //{
        //    return (value is Byte ||
        //            value is Int16 ||
        //            value is Int32 ||
        //            value is Int64 ||
        //            value is SByte ||
        //            value is UInt16 ||
        //            value is UInt32 ||
        //            value is UInt64 ||
        //            value is BigInteger ||
        //            value is Decimal ||
        //            value is Double ||
        //            value is Single);
        //}

        public static bool IsNumeric(System.Object value)
        {
            return (value is Byte ||
                    value is Int16 ||
                    value is Int32 ||
                    value is Int64 ||
                    value is SByte ||
                    value is UInt16 ||
                    value is UInt32 ||
                    value is UInt64 ||
                    value is BigInteger || //需要引入System.Numerics
                    value is Decimal ||
                    value is Double ||
                    value is Single);
        }

    }
}