using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using System.Collections;

namespace AuroraZhou.support
{
    /// <summary>
    /// 最后更新日期2018年8月30日
    /// 修正bug 463行：错误的合并单元格判断
    /// </summary>
    public class NpoiHelperV1
    {
        /// <summary>
        /// 字符型单元格外接处理委托
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public delegate string StringProcessing(object obj);
        /// <summary>
        /// 单元格更新事件
        /// </summary>
        public static event StringProcessing CellFormulaChangeEvent;

        public static Exception exception;
        /// <summary>
        /// 同一个sheet的行复制，
        /// 注意：复制的行中存在夸行合并单元格时将会出错，请调用CopyRowWithoutMergedRegion
        /// </summary>
        /// <param name="workbook">将要更改的iworkbook</param>
        /// <param name="worksheet">要复制行的isheet</param>
        /// <param name="sourceRowNum">源行号</param>
        /// <param name="destinationRowNum">目标行号</param>
        public static void CopyRow(IWorkbook workbook, ISheet worksheet, int sourceRowNum, int destinationRowNum)
        {
            // Get the source / new row
            IRow newRow = worksheet.GetRow(destinationRowNum);
            IRow sourceRow = worksheet.GetRow(sourceRowNum);
            // If the row exist in destination, push down all rows by 1 else create a new row
            if (newRow != null)
            {
                worksheet.ShiftRows(destinationRowNum, worksheet.LastRowNum, 1, true, false);
            }
            else
            {
                newRow = worksheet.CreateRow(destinationRowNum);
            }
            newRow.Height = sourceRow.Height;
            // 遍历一次旧行并添加到新行
            for (int i = 0; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                ICell oldCell = sourceRow.GetCell(i);
                ICell newCell = newRow.CreateCell(i);

                // 如果旧行为空则跳过
                if (oldCell == null)
                {
                    newCell = null;
                    continue;
                }
                newCell.CellStyle = oldCell.CellStyle;

                // 复制批注
                if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

                // 复制超链
                if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

                // 复制单元格类型
                newCell.SetCellType(oldCell.CellType);

                // 设置单元格数据值
                switch (oldCell.CellType)
                {
                    case CellType.Blank://空值
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                    case CellType.Boolean:
                        newCell.SetCellValue(oldCell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                        break;
                    case CellType.Formula://公式
                        newCell.SetCellFormula(oldCell.CellFormula);
                        break;
                    case CellType.Numeric://数值
                        newCell.SetCellValue(oldCell.NumericCellValue);
                        break;
                    case CellType.String://字符串
                        newCell.SetCellValue(oldCell.RichStringCellValue);
                        break;
                    case CellType.Unknown://未知
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                }
            }

            // 合拼单元格操作
            for (int i = 0; i < worksheet.NumMergedRegions; i++)//NumMergedRegions：整个sheet的合拼单元格数量
            {
                NPOI.SS.Util.CellRangeAddress cellRangeAddress = worksheet.GetMergedRegion(i);//获取合拼单元格的地址字符串
                if (cellRangeAddress.FirstRow == sourceRow.RowNum)
                {
                    NPOI.SS.Util.CellRangeAddress newCellRangeAddress = new NPOI.SS.Util.CellRangeAddress(newRow.RowNum,
                                                                                (newRow.RowNum +
                                                                                 (cellRangeAddress.FirstRow -
                                                                                  cellRangeAddress.LastRow)),
                                                                                cellRangeAddress.FirstColumn,
                                                                                cellRangeAddress.LastColumn);
                    worksheet.AddMergedRegion(newCellRangeAddress);//跨行单元格合拼时发生错误
                }
            }
        }

        /// <summary>
        /// 跨sheet复制行
        /// 注意：复制的行中存在夸行合并单元格时将会出错
        /// </summary>
        /// <param name="sourceSheet">源表</param>
        /// <param name="targetSheet">目标表</param>
        /// <param name="sourceRowNum">源复制行</param>
        /// <param name="targetRowNum">目标复制行</param>
        /// <param name="doMerged">是否复制合并单元格</param>
        ///<param name="cover">是否以覆盖原来内容，true为覆盖，false为插入</param> 
        public static void CopyRowOverSheet(ISheet sourceSheet, ISheet targetSheet, int sourceRowNum, int targetRowNum, bool doMerged, bool cover)
        {
            IRow targetRow = targetSheet.GetRow(targetRowNum);
            IRow sourceRow = sourceSheet.GetRow(sourceRowNum);
            if (targetRow != null && !cover)
            {
                //如果插入的行已经有内容，则将该内容以及该行以下的全部内容都往下移一行
                targetSheet.ShiftRows(targetRowNum, targetSheet.LastRowNum, 1, true, false);
            }
            else
            {
                targetRow = targetSheet.CreateRow(targetRowNum);
            }
            short h = sourceRow.Height;
            targetRow.Height = h;

            for (int i = 0; i < sourceRow.LastCellNum; i++)
            {
                ICell oldCell = sourceRow.GetCell(i);
                ICell newCell = targetRow.CreateCell(i);
                if (oldCell == null)
                {
                    newCell = null;
                    continue;
                }
                if (oldCell.CellStyle != null)
                {
                    ICellStyle os = oldCell.CellStyle;
                    if (sourceSheet.Workbook == targetSheet.Workbook)
                    {
                        newCell.CellStyle = os;
                    }
                    else
                    {//跨workbook的情况
                        newCell.CellStyle = cloneCellstyle(sourceSheet.Workbook, targetSheet.Workbook, os);
                    }
                }
                // 复制批注
                if (oldCell.CellComment != null)
                {
                    IComment ic = oldCell.CellComment;
                    newCell.CellComment = ic;
                }
                // 复制超链
                if (oldCell.Hyperlink != null)
                {
                    IHyperlink hl = oldCell.Hyperlink;
                    newCell.Hyperlink = hl;
                }
                // 复制单元格类型
                CellType ct = oldCell.CellType;
                newCell.SetCellType(ct);
                // 设置单元格数据值
                switch (oldCell.CellType)
                {
                    case CellType.Blank://空值
                        string s1 = oldCell.StringCellValue;
                        newCell.SetCellValue(s1);
                        break;
                    case CellType.Boolean:
                        bool b1 = oldCell.BooleanCellValue;
                        newCell.SetCellValue(b1);
                        break;
                    case CellType.Error:
                        byte err = oldCell.ErrorCellValue;
                        newCell.SetCellErrorValue(err);
                        break;
                    case CellType.Formula://公式,重要更新公式可以外接处理程序
                        string f = oldCell.CellFormula;
                        if (CellFormulaChangeEvent != null)
                        {
                            Tuple<string, ICell, ICell> t = new Tuple<string, ICell, ICell>(f, oldCell, newCell);
                            f = CellFormulaChangeEvent(t);
                        }
                        newCell.SetCellFormula(f);
                        break;
                    case CellType.Numeric://数值
                        double d = oldCell.NumericCellValue;
                        newCell.SetCellValue(d);
                        break;
                    case CellType.String://字符串
                        IRichTextString ir = oldCell.RichStringCellValue;
                        newCell.SetCellValue(ir);
                        break;
                    case CellType.Unknown://未知
                        string scv = oldCell.StringCellValue;
                        newCell.SetCellValue(scv);
                        break;
                }
            }

            // 合拼单元格操作
            if (doMerged)
            {
                for (int i = 0; i < sourceSheet.NumMergedRegions; i++)//NumMergedRegions：整个sheet的合拼单元格数量
                {
                    NPOI.SS.Util.CellRangeAddress cellRangeAddress = sourceSheet.GetMergedRegion(i);//获取合拼单元格的地址字符串
                    if (cellRangeAddress.FirstRow == sourceRow.RowNum)
                    {
                        NPOI.SS.Util.CellRangeAddress newCellRangeAddress = new NPOI.SS.Util.CellRangeAddress(targetRow.RowNum,
                                                                                    (targetRow.RowNum +
                                                                                     (cellRangeAddress.FirstRow -
                                                                                      cellRangeAddress.LastRow)),
                                                                                    cellRangeAddress.FirstColumn,
                                                                                    cellRangeAddress.LastColumn);
                        targetSheet.AddMergedRegion(newCellRangeAddress);
                    }
                }
            }
        }

        /// <summary>
        /// 跨workbook的cellstyle复制
        /// 注意：只在xls生效，且只能复制预设颜色，其他颜色会被替换
        /// </summary>
        /// <param name="iWorkbook">源workbook</param>
        /// <param name="iWorkbook_2">目标workbook</param>
        /// <param name="os"></param>
        /// <returns></returns>
        private static ICellStyle cloneCellstyle(IWorkbook sourceWorkbook, IWorkbook targetWorkbook, ICellStyle sourceCellStyle)
        {
            ICellStyle cs = targetWorkbook.CreateCellStyle();
            cs.Alignment = sourceCellStyle.Alignment;
            cs.BorderBottom = sourceCellStyle.BorderBottom;
            cs.BorderDiagonal = sourceCellStyle.BorderDiagonal;
            cs.BorderDiagonalColor = sourceCellStyle.BorderDiagonalColor;
            cs.BorderDiagonalLineStyle = sourceCellStyle.BorderDiagonalLineStyle;
            cs.BorderLeft = sourceCellStyle.BorderLeft;
            cs.BorderRight = sourceCellStyle.BorderRight;
            cs.BorderTop = sourceCellStyle.BorderTop;
            cs.BottomBorderColor = sourceCellStyle.BottomBorderColor;
            cs.DataFormat = sourceCellStyle.DataFormat;
            //只有在xls的workbook下才能设置仅有的颜色，其他颜色会被丢弃
            //而在xlsx中FillForegroundColor的index都为0，所以导致复制的单元格全部为黑色背景色
            cs.FillBackgroundColor = sourceCellStyle.FillBackgroundColor;
            cs.FillForegroundColor = sourceCellStyle.FillForegroundColor;//设置背景色
            cs.FillPattern = sourceCellStyle.FillPattern;//填充样式
            cs.Indention = sourceCellStyle.Indention;//缩进
            cs.IsHidden = sourceCellStyle.IsHidden;//单元格隐藏
            cs.IsLocked = sourceCellStyle.IsLocked;//单元格锁定
            cs.LeftBorderColor = sourceCellStyle.LeftBorderColor;
            cs.RightBorderColor = sourceCellStyle.RightBorderColor;
            cs.Rotation = sourceCellStyle.Rotation;//旋转
            cs.ShrinkToFit = sourceCellStyle.ShrinkToFit;//自动匹配大小
            cs.TopBorderColor = sourceCellStyle.TopBorderColor;
            cs.VerticalAlignment = sourceCellStyle.VerticalAlignment;//垂直对齐
            cs.WrapText = sourceCellStyle.WrapText;//自动换行
            return cs;
        }

        /// <summary>
        /// Cellstyle缓存
        /// </summary>
        public class CellStyleCache : ArrayList
        {
            public ICellStyle this[ICellStyle fromStyle]
            {
                get
                {
                    foreach (object o in this)
                    {
                        ICellStyle toStyle = o as ICellStyle;
                        if (//以下列举cellstyle已知的全部属性
                            toStyle.Alignment == fromStyle.Alignment//对齐
                            && toStyle.BorderBottom == fromStyle.BorderBottom
                            && toStyle.BorderDiagonal == fromStyle.BorderDiagonal
                            && toStyle.BorderDiagonalColor == fromStyle.BorderDiagonalColor
                            && toStyle.BorderDiagonalLineStyle == fromStyle.BorderDiagonalLineStyle
                            && toStyle.BorderLeft == fromStyle.BorderLeft
                            && toStyle.BorderRight == fromStyle.BorderRight
                            && toStyle.BorderTop == fromStyle.BorderTop
                            && toStyle.BottomBorderColor == fromStyle.BottomBorderColor
                            && toStyle.DataFormat == fromStyle.DataFormat
                            && toStyle.FillBackgroundColor == fromStyle.FillBackgroundColor
                            && toStyle.FillBackgroundColorColor == fromStyle.FillBackgroundColorColor
                            && toStyle.FillForegroundColor == fromStyle.FillForegroundColor
                            && toStyle.FillForegroundColorColor == fromStyle.FillForegroundColorColor
                            && toStyle.FillPattern == fromStyle.FillPattern//填充样式
                            && toStyle.FontIndex == fromStyle.FontIndex
                            && toStyle.Indention == fromStyle.Indention//缩进
                            //&& toStyle.Index == fromStyle.Index//未知用途暂不判断
                            && toStyle.IsHidden == fromStyle.IsHidden//单元格隐藏
                            && toStyle.IsLocked == fromStyle.IsLocked//单元格锁定
                            && toStyle.LeftBorderColor == fromStyle.LeftBorderColor
                            && toStyle.RightBorderColor == fromStyle.RightBorderColor
                            && toStyle.Rotation == fromStyle.Rotation//旋转
                            && toStyle.ShrinkToFit == fromStyle.ShrinkToFit//自动匹配大小
                            && toStyle.TopBorderColor == fromStyle.TopBorderColor
                            && toStyle.VerticalAlignment == fromStyle.VerticalAlignment//垂直对齐
                            && toStyle.WrapText == fromStyle.WrapText//自动换行
                            )
                        { return toStyle; }
                    }
                    return null;
                }
                set { this.Add(fromStyle); }
            }
        }

        /// <summary>
        /// 同一个sheet的行复制，但不合拼单元格
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="worksheet"></param>
        /// <param name="sourceRowNum"></param>
        /// <param name="destinationRowNum"></param>
        public static void CopyRowWithoutMergedRegion(IWorkbook workbook, ISheet worksheet, int sourceRowNum, int destinationRowNum)
        {
            // Get the source / new row
            IRow newRow = worksheet.GetRow(destinationRowNum);
            IRow sourceRow = worksheet.GetRow(sourceRowNum);
            // If the row exist in destination, push down all rows by 1 else create a new row
            if (newRow != null)
            {
                worksheet.ShiftRows(destinationRowNum, worksheet.LastRowNum, 1, true, false);
            }
            else
            {
                newRow = worksheet.CreateRow(destinationRowNum);
            }
            //设置行高
            newRow.Height = sourceRow.Height;
            //设置分隔行
            if (worksheet.IsRowBroken(sourceRowNum))
            {
                worksheet.SetRowBreak(destinationRowNum);
            }

            // 遍历一次旧行并添加到新行
            for (int i = 0; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                ICell oldCell = sourceRow.GetCell(i);
                ICell newCell = newRow.CreateCell(i);

                // 如果旧行为空则跳过
                if (oldCell == null)
                {
                    newCell = null;
                    continue;
                }
                newCell.CellStyle = oldCell.CellStyle;

                // 复制批注
                if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

                // 复制超链
                if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

                // 复制单元格类型
                newCell.SetCellType(oldCell.CellType);

                // 设置单元格数据值
                switch (oldCell.CellType)
                {
                    case CellType.Blank://空值
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                    case CellType.Boolean:
                        newCell.SetCellValue(oldCell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                        break;
                    case CellType.Formula://公式
                        newCell.SetCellFormula(oldCell.CellFormula);
                        break;
                    case CellType.Numeric://数值
                        newCell.SetCellValue(oldCell.NumericCellValue);
                        break;
                    case CellType.String://字符串
                        newCell.SetCellValue(oldCell.RichStringCellValue);
                        break;
                    case CellType.Unknown://未知
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                }
            }
        }

        /// <summary>
        /// 复制多行并合拼其中的单元格
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="worksheet">目标sheet</param>
        /// <param name="sourceRowNum">起始行号</param>
        /// <param name="copyRowsCount">复制行数</param>
        /// <param name="destinationRowNum">目标行号</param>
        public static void CopyRows(IWorkbook workbook, ISheet worksheet, int sourceRowNum, int copyRowsCount, int destinationRowNum)
        {
            int[,] sm = getAllMergedRegions(worksheet);
            List<int[]> copyAreaMergedRegions = new List<int[]>();
            for (int i = 0; i < sm.GetLength(0); i++)
            {
                if (sm[i, 0] >= sourceRowNum && sm[i, 0] < (sourceRowNum + copyRowsCount))
                {
                    copyAreaMergedRegions.Add(new int[4] { sm[i, 0] - sourceRowNum, sm[i, 1], sm[i, 2] - sourceRowNum, sm[i, 3] });
                }
            }
            int addpoint = destinationRowNum;
            for (int row = 0; row < copyRowsCount; row++)
            {
                //复制行
                NpoiHelperV1.CopyRowWithoutMergedRegion(workbook, worksheet, sourceRowNum + row, addpoint);
                addpoint++;
            }
            foreach (int[] cell in copyAreaMergedRegions)
            {
                NPOI.SS.Util.CellRangeAddress newCellRangeAddress = new NPOI.SS.Util.CellRangeAddress(
                    cell[0] + destinationRowNum, //顶
                    cell[2] + destinationRowNum, //底
                    cell[1], //左
                    cell[3]  //右
                    );
                worksheet.AddMergedRegion(newCellRangeAddress);
            }
        }

        /// <summary>
        /// 跨sheet复制多行，连合并单元格
        /// </summary>
        /// <param name="worksheet">源sheet</param>
        /// <param name="targetsheet">目标</param>
        /// <param name="sourceRowNum">源行值</param>
        /// <param name="copyRowsCount">复制行数</param>
        /// <param name="destinationRowNum">目标行值</param>
        /// <param name="cover">是否以覆盖原来内容，true为覆盖，false为插入</param>
        public static void CopyRowsOverSheet(ISheet worksheet, ISheet targetsheet,
            int sourceRowNum, int copyRowsCount, int destinationRowNum, bool cover)
        {
            int[,] sm = getAllMergedRegions(worksheet);
            List<int[]> copyAreaMergedRegions = new List<int[]>();
            //遍历一次源sheet的全部合并单元格
            for (int i = 0; i < sm.GetLength(0); i++)
            {
                if (sm[i, 0] >= sourceRowNum && sm[i, 0] < (sourceRowNum + copyRowsCount))//18年8月30日发现bug：小于等于改为小于
                {
                    copyAreaMergedRegions.Add(new int[4] { sm[i, 0] - sourceRowNum, sm[i, 1], sm[i, 2] - sourceRowNum, sm[i, 3] });
                }
            }
            //int addpoint = destinationRowNum;
            for (int row = 0; row < copyRowsCount; row++)
            {
                //跨sheet复制行
                NpoiHelperV1.CopyRowOverSheet(worksheet, targetsheet, sourceRowNum + row, destinationRowNum + row, false, cover);
            }
            //将新复制的行进行合并
            foreach (int[] cell in copyAreaMergedRegions)
            {
                NPOI.SS.Util.CellRangeAddress newCellRangeAddress = new NPOI.SS.Util.CellRangeAddress(
                    cell[0] + destinationRowNum, //顶
                    cell[2] + destinationRowNum, //底
                    cell[1], //左
                    cell[3]  //右
                    );
                targetsheet.AddMergedRegion(newCellRangeAddress);
            }
        }

        /// <summary>
        /// 删除行，解决因为shiftRows上移删除行而造成的格式错乱问题
        /// </summary>
        /// <param name="worksheet">目标sheet</param>
        /// <param name="startRow">需要删除开始行(从0开始)</param>
        /// <param name="count">删除行数</param>
        public static void DelRows(ISheet worksheet, int startRow, int count)
        {
            int NumMergedRegions = worksheet.NumMergedRegions;
            //int[,] allMergedRegion = getAllMergedRegions(worksheet);
            for (int i = NumMergedRegions - 1; i >= 0; i--)
            {
                NPOI.SS.Util.CellRangeAddress cellRangeAddress = worksheet.GetMergedRegion(i);
                if (cellRangeAddress.FirstRow >= startRow + count
                    || cellRangeAddress.LastRow <= startRow)
                {
                    //只有一行的合并单元格FirstRow==LastRow
                    if (cellRangeAddress.FirstRow == cellRangeAddress.LastRow)
                    {
                        //刚好在删除区域的startRow或endRow
                        if (cellRangeAddress.FirstRow == startRow
                            || cellRangeAddress.FirstRow == startRow + count - 1)
                        {
                            worksheet.RemoveMergedRegion(i);
                        }
                    }
                }
                else
                {
                    //该合并单元格的行已经在被删除的区域内，所以解除该合并单元格
                    //如果不删除该合并单元格将会导致全部格式错乱
                    worksheet.RemoveMergedRegion(i);
                }
            }
            worksheet.ShiftRows(startRow + count, worksheet.LastRowNum, -count, true, false);
        }

        /// <summary>
        /// 获取一个sheet的所有合并单元格
        /// </summary>
        /// <param name="worksheet">目标sheet</param>
        /// <returns>返回数组结构：起始行，起始列，终止行，终止列；所以第一二个值可用于对该合并格赋值或读取
        /// 但该数组的排序根据操作的顺序而定，最新加入的合并单元格为第0个
        /// </returns>
        public static int[,] getAllMergedRegions(ISheet worksheet)
        {
            int NumMergedRegions = worksheet.NumMergedRegions;
            int[,] output = new int[NumMergedRegions, 4];
            for (int i = 0; i < worksheet.NumMergedRegions; i++)//NumMergedRegions：整个sheet的合拼单元格数量
            {
                NPOI.SS.Util.CellRangeAddress cellRangeAddress = worksheet.GetMergedRegion(i);//获取合拼单元格的地址字符串
                output[i, 0] = cellRangeAddress.FirstRow;
                output[i, 1] = cellRangeAddress.FirstColumn;
                output[i, 2] = cellRangeAddress.LastRow;
                output[i, 3] = cellRangeAddress.LastColumn;
            }
            return output;
        }

        /// <summary>
        /// 设置所有sheet自动计算
        /// </summary>
        public static void setAllSheetToAuto(IWorkbook loadedExcel)
        {
            if (loadedExcel.NumberOfSheets > 0)
            {
                int x = loadedExcel.NumberOfSheets;
                for (int i = 0; i < x; i++)
                {
                    loadedExcel.GetSheetAt(i).ForceFormulaRecalculation = true; //让公式自动计算
                }
            }
        }

    }
}
