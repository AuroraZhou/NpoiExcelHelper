using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.Collections;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;

namespace AuroraZhou
{
	/// <summary>
    /// excel地址相关处理及公式升级
    /// </summary>
    public class ExcelAddress
    {
        static ExcelAddress()
        {

        }

        /// <summary>
        /// 将指定的自然数转换为26进制表示。映射关系：[1-26] ->[A-Z]。
        /// </summary>
        /// <param name="n">自然数（如果无效，则返回空字符串）。</param>
        /// <returns>26进制表示。</returns>
        public static string ToNumberSystem26(int input)
        {
            int n = input + 1;
            string s = string.Empty;
            while (n > 0)
            {
                int m = n % 26;
                if (m == 0) m = 26;
                s = (char)(m + 64) + s;
                n = (n - m) / 26;
            }
            return s;
        }

        /// <summary>
        /// 将指定的26进制表示转换为自然数。映射关系：[A-Z] ->[1-26]。
        /// </summary>
        /// <param name="s">26进制表示（如果无效，则返回0）。</param>
        /// <returns>自然数。</returns>
        public static int FromNumberSystem26(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            int n = 0;
            for (int i = s.Length - 1, j = 1; i >= 0; i--, j *= 26)
            {
                char c = Char.ToUpper(s[i]);
                if (c < 'A' || c > 'Z') return 0;
                n += ((int)c - 64) * j;
            }
            return n;
        }

        /// <summary>
        /// 从地址翻译为行列
        /// </summary>
        /// <param name="str">地址字符串</param>
        /// <returns>1为行，2为列</returns>
        public static Tuple<int, int> GetRowColNumber(string str)
        {
            string expr = "[a-zA-Z]{1,3}";
            string n = "[0-9]{1,7}";
            if (str == null) return null;
            MatchCollection az = Regex.Matches(str, expr);
            MatchCollection num = Regex.Matches(str, n);
            int[] output = new int[2];
            output[0] = int.Parse(num[0].Value) - 1;//行
            output[1] = FromNumberSystem26(az[0].Value) - 1;//列
            return new Tuple<int, int>(output[0], output[1]);
        }

        /// <summary>
        /// 求出两个单元格的差(用于公式更新)
        /// </summary>
        /// <param name="v1">单元格内的公式引用地址</param>
        /// <param name="v2">单元格自身的地址</param>
        /// <returns>相差的行与列</returns>
        public static Tuple<int, int> Subtraction(string v1, string v2)
        {
            Tuple<int, int> t1 = GetRowColNumber(v1);
            Tuple<int, int> t2 = GetRowColNumber(v2);
            return new Tuple<int, int>(t1.Item1 - t2.Item1, t1.Item2 - t2.Item2);
        }

        /// <summary>
        /// 求出单元格的映射目标(用于公式更新)
        /// </summary>
        /// <param name="v">单元格自身的地址</param>
        /// <param name="c">相差的行与列</param>
        /// <returns>单元格内的公式引用地址</returns>
        public static string Addition(string v, Tuple<int, int> c)
        {
            Tuple<int, int> o = GetRowColNumber(v);
            return GetCellAddress(new Tuple<int, int>(o.Item1 + c.Item1, o.Item2 + c.Item2));
        }

        /// <summary>
        /// 行列号转字符串地址
        /// </summary>
        /// <param name="rc"></param>
        /// <returns></returns>
        public static string GetCellAddress(Tuple<int, int> rc)
        {
            //起始行列都是从0开始，所以需要加一对应字典与地址
            string r = (rc.Item1 + 1).ToString();
            return ToNumberSystem26(rc.Item2) + r;
        }

        /// <summary>
        /// 从单元格获得字符串地址
        /// </summary>
        /// <param name="ic">单元格</param>
        /// <returns>字符串地址</returns>
        public static string GetCellAddress(ICell ic)
        {
            Tuple<int, int> t1 = new Tuple<int, int>(ic.RowIndex, ic.ColumnIndex);
            return GetCellAddress(t1);
        }

        /// <summary>
        /// 从单元格获得字符串地址
        /// </summary>
        /// <param name="row">行索引号</param>
        /// <param name="col">列索引号</param>
        /// <returns></returns>
        public static string GetCellAddress(int row, int col)
        {
            Tuple<int, int> t1 = new Tuple<int, int>(row, col);
            return GetCellAddress(t1);
        }

        /// <summary>
        /// 更新单元格公式，目前未支持特殊符号的处理，例如$
        /// </summary>
        /// <param name="formula">原单元格内公式</param>
        /// <param name="sourceCellAdd">原单元格地址</param>
        /// <param name="targetCellAdd">目标单元格地址</param>
        /// <returns>更新后公式</returns>
        public static string UpdateFormula(string formula, string sourceCellAdd, string targetCellAdd)
        {
            string expr = "[a-zA-Z]{1,3}[0-9]{1,7}";//excel最大单元格为XFD1048576
            MatchCollection result = Regex.Matches(formula, expr);
            string output = formula;
            if (result.Count > 0)
            {
                for (int i = 0; i < result.Count; i++)
                {
                    Tuple<int, int> cell2cell = Subtraction(result[i].Value, sourceCellAdd);
                    output = output.Replace(result[i].Value, Addition(targetCellAdd, cell2cell));
                }
            }
            return output;
        }

        /// <summary>
        /// 更新单元格公式，目前未支持特殊符号的处理，例如$
        /// </summary>
        /// <param name="formula">原单元格内公式</param>
        /// <param name="source">源单元格</param>
        /// <param name="target">目标单元格</param>
        /// <returns></returns>
        public static string UpdateFormula(string formula, ICell source, ICell target)
        {
            string s = GetCellAddress(source);
            string t = GetCellAddress(target);
            return UpdateFormula(formula, s, t);
        }
    }
}