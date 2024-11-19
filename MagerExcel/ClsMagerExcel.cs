using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;

namespace MagerExcel
{
    public class ClsMagerExcel
    {
        public class CellResult
        {
            public int Row { get; set; }
            public int Col { get; set; }
        }

        public ExcelWorksheet ws;
        public int rowStart;
        public int colStart;

        public ClsMagerExcel(ExcelWorksheet _ws)
        {
            ws = _ws;
            rowStart = 0;
            colStart = 0;
        }

        public ClsMagerExcel(ExcelWorksheet _ws, double columnWidth = 8.43)
        {
            ws = _ws;
            rowStart = 0;
            colStart = 0;

            ws.DefaultColWidth = columnWidth;
        }

        public ClsMagerExcel(ExcelWorksheet _ws, double columnWidth = 8.43, int rowHeight = 15)
        {
            ws = _ws;
            rowStart = 0;
            colStart = 0;

            ws.DefaultColWidth = columnWidth;
            ws.DefaultRowHeight = rowHeight;
        }

        public enum BorderType
        {
            NoBorderAll,
            NoBorderAround,
            BorderAroundThin,
            BorderAroundThick,
            BorderAllThin,
            BorderAllThick,
            BorderAroundDotted,
            BorderAroundDashed,
            BorderAllDotted,
            BorderAllDashed,
            BorderBottomThin,
            BorderBottomThick,
            BorderBottomDotted,
            BorderBottomDashed,
            BorderBottomDouble,
            BorderAllDouble,
            BorderAroundDouble
        }

        public enum BgColorType
        {
            Odd,
            Even,
            All
        }

        //-------------------------------------------------------------- ROW / COLUMN SETTING -----------------------------------------------------------------------

        public void SetColWidth(int col, double width)
        {
            ws.Column(col).Width = width;
        }

        public void SetColWidthByRange(int col, int colTo, double width)
        {
            while (col <= colTo)
            {
                ws.Column(col).Width = width;
                col++;
            }
        }

        public void SetRowHeight(int row, double height)
        {
            ws.Row(row).Height = height;
        }

        public void SetRowHeightByRange(int row, int rowTo, double height)
        {
            while (row <= rowTo)
            {
                ws.Row(row).Height = height;
                row++;
            }
        }

        //-------------------------------------------------------------- BOLD ONLY ----------------------------------------------------------------------------------

        public CellResult Bold(int row, int col)
        {
            try
            {
                ws.Cells[row, col].Style.Font.Bold = true;
                return new CellResult { Row = row, Col = col };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[Bold] row({row}) column({col})");
            }
        }

        public CellResult Bold(int row, int col, int rowTo, int colTo)
        {
            try
            {
                ws.Cells[row, col, rowTo, colTo].Style.Font.Bold = true;
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[Bold] row({row}) column({col})");
            }
        }
        //-------------------------------------------------------------- MERGE ONLY ----------------------------------------------------------------------------------

        public CellResult Merge(int row, int col, int rowTo, int colTo)
        {
            try
            {
                ws.Cells[row, col, rowTo, colTo].Merge = true;
                ws.Cells[row, col, rowTo, colTo].Style.WrapText = true;
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[Merge] row({row}) column({col})");
            }
        }

        //-------------------------------------------------------------- ALIGNMENT ----------------------------------------------------------------------------------
        public CellResult SetAlign(int row, int col, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                ws.Cells[row, col].Style.VerticalAlignment = vAlign;
                ws.Cells[row, col].Style.HorizontalAlignment = hAlign;
                return new CellResult { Row = row, Col = col };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[SetAlign] row({row}) column({col})");
            }
        }

        public CellResult SetAlign(int row, int col, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign, int rowTo, int colTo)
        {
            try
            {
                ws.Cells[row, col, rowTo, colTo].Style.VerticalAlignment = vAlign;
                ws.Cells[row, col, rowTo, colTo].Style.HorizontalAlignment = hAlign;
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[SetAlign] row({row}) column({col})");
            }
        }

        //-------------------------------------------------------------- TRANSPOSE / TABLE RIGHT -------------------------------------------------------------------

        public CellResult DrawTableRightMerge(int row, int col, List<object> dataList)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    rowStart = row;
                    int rowLast = 0;

                    foreach (object data in dataList)
                    {
                        Type dataType = data.GetType();
                        PropertyInfo[] properties = dataType.GetProperties();

                        foreach (PropertyInfo property in properties)
                        {
                            object propertyValue = property.GetValue(data);
                            ws.Cells[row, col].Value = propertyValue;
                            ws.Cells[row, col].Style.WrapText = true;
                            row++;
                        }
                        col++;
                        rowLast = row > rowLast ? row : rowLast;
                        row = rowStart;
                    }
                    return new CellResult { Row = rowLast, Col = col };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawTableRightMerge] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawTableRightMerge] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawTableRightMerge(int row, int col, List<object> dataList, int rowMerge, int colMerge)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    rowMerge = rowMerge < 1 ? 1 : rowMerge;
                    colMerge = colMerge < 1 ? 1 : colMerge;
                    rowStart = row;
                    colStart = col;
                    int rowLast = 0;

                    foreach (object data in dataList)
                    {
                        Type dataType = data.GetType();
                        PropertyInfo[] properties = dataType.GetProperties();

                        foreach (PropertyInfo property in properties)
                        {
                            object propertyValue = property.GetValue(data);
                            ws.Cells[row, col].Value = propertyValue;
                            ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Merge = true;
                            ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Style.WrapText = true;
                            row += rowMerge;
                        }
                        col += colMerge;
                        rowLast = row > rowLast ? row : rowLast;
                        row = rowStart;
                    }
                    return new CellResult { Row = rowLast, Col = col };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawTableRightMerge] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawTableRightMerge] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawTableRightMerge(int row, int col, List<object> dataList, int rowMerge, int colMerge, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                CellResult result = DrawTableRightMerge(row, col, dataList, rowMerge, colMerge);
                return SetAlign(row, col, vAlign, hAlign, result.Row, result.Col);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawTableRightMerge] row({row}) column({col})");
            }
        }

        public CellResult DrawTableRightMerge(int row, int col, List<object> dataList, int rowMerge, int colMerge, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign, BorderType borderType)
        {
            try
            {
                CellResult result = DrawTableRightMerge(row, col, dataList, rowMerge, colMerge);
                result = SetAlign(row, col, vAlign, hAlign, result.Row, result.Col);
                return DrawBorderByType(row, col, result.Row - 1, result.Col - 1, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawTableRightMerge] row({row}) column({col})");
            }
        }

        //-------------------------------------------------------------- HEADER RIGHT -------------------------------------------------------------------------------

        public CellResult DrawHeaderRight(int row, int col, object obj)
        {
            if (obj != null)
            {
                try
                {
                    Type type = obj.GetType();
                    PropertyInfo[] properties = type.GetProperties();

                    foreach (PropertyInfo property in properties)
                    {
                        ws.Cells[row, col].Value = property.Name; col++;
                    }
                    return new CellResult { Row = row, Col = col };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawHeaderRight] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawHeaderRight] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawHeaderRight(int row, int col, object obj, int rowMerge, int colMerge)
        {
            if (obj != null)
            {
                try
                {
                    rowStart = row;
                    colStart = col;
                    rowMerge = rowMerge < 1 ? 1 : rowMerge;
                    colMerge = colMerge < 1 ? 1 : colMerge;

                    Type type = obj.GetType();
                    PropertyInfo[] properties = type.GetProperties();

                    foreach (PropertyInfo property in properties)
                    {
                        ws.Cells[row, col].Value = property.Name;
                        ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Merge = true;
                        ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Style.WrapText = true;
                        col += colMerge;
                    }
                    return new CellResult { Row = row, Col = col };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawHeaderRight] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawHeaderRight] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawHeaderRight(int row, int col, object obj, int rowMerge, int colMerge, BorderType borderType)
        {
            try
            {
                CellResult result = DrawHeaderRight(row, col, obj, rowMerge, colMerge);
                return DrawBorderByType(rowStart, colStart, result.Row + rowMerge - 1, result.Col - 1, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawHeaderRight] row({row}) column({col})");
            }
        }

        public CellResult DrawHeaderRight(int row, int col, object obj, int rowMerge, int colMerge, BorderType borderType, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                CellResult result = DrawHeaderRight(row, col, obj, rowMerge, colMerge, borderType);
                return SetAlign(rowStart, colStart, vAlign, hAlign, result.Row + rowMerge, result.Col - 1);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawHeaderRight] row({row}) column({col})");
            }
        }

        //-------------------------------------------------------------- HEADER DOWN -------------------------------------------------------------------------------

        public CellResult DrawHeaderDown(int row, int col, object obj)
        {
            if (obj != null)
            {
                try
                {
                    Type type = obj.GetType();
                    PropertyInfo[] properties = type.GetProperties();

                    foreach (PropertyInfo property in properties)
                    {
                        ws.Cells[row, col].Value = property.Name; row++;
                    }
                    return new CellResult { Row = row, Col = col };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawHeaderDown] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawHeaderDown] object/list is NULL. row({row}) column({col})");
            }
        }

        //-------------------------------------------------------------- LIST DOWN -------------------------------------------------------------------------------

        public CellResult DrawListDown(int row, int col, List<object> dataList)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    foreach (var obj in dataList)
                    {
                        ws.Cells[row, col].Value = obj; row++;
                    }
                    return new CellResult { Row = row, Col = col };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawListDown] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawListDown] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawListDown(int row, int col, List<object> dataList, int rowMerge, int colMerge)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    rowStart = row;
                    colStart = col;
                    rowMerge = rowMerge < 1 ? 1 : rowMerge;
                    colMerge = colMerge < 1 ? 1 : colMerge;

                    foreach (var obj in dataList)
                    {
                        ws.Cells[row, col].Value = obj;
                        ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Merge = true;
                        ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Style.WrapText = true;
                        row += rowMerge;
                        col = colStart;
                    }
                    return new CellResult { Row = row, Col = col + colMerge };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawListDown] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawListDown] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawListDown(int row, int col, List<object> dataList, int rowMerge, int colMerge, BorderType borderType)
        {
            try
            {
                CellResult result = DrawListDown(row, col, dataList, rowMerge, colMerge);
                return DrawBorderByType(rowStart, colStart, result.Row - 1, result.Col - 1, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawListDown] row({row}) column({col})");
            }
        }

        public CellResult DrawListDown(int row, int col, List<object> dataList, int rowMerge, int colMerge, BorderType borderType, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                CellResult result = DrawListDown(row, col, dataList, rowMerge, colMerge, borderType);
                return SetAlign(rowStart, colStart, vAlign, hAlign, result.Row, result.Col);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawListDown] row({row}) column({col})");
            }
        }

        //-------------------------------------------------------------- LIST RIGHT -------------------------------------------------------------------------------

        public CellResult DrawListRight(int row, int col, List<object> dataList)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    foreach (var obj in dataList)
                    {
                        ws.Cells[row, col].Value = obj; col++;
                    }
                    return new CellResult { Row = row, Col = col };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawListRight] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawListRight] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawListRight(int row, int col, List<object> dataList, int rowMerge, int colMerge)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    rowStart = row;
                    colStart = col;
                    rowMerge = rowMerge < 1 ? 1 : rowMerge;
                    colMerge = colMerge < 1 ? 1 : colMerge;

                    foreach (var obj in dataList)
                    {
                        ws.Cells[row, col].Value = obj;
                        ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Merge = true;
                        ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Style.WrapText = true;
                        col += colMerge;
                    }
                    return new CellResult { Row = rowStart, Col = col + colMerge };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawListRight] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawListRight] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawListRight(int row, int col, List<object> dataList, int rowMerge, int colMerge, BorderType borderType)
        {
            try
            {
                CellResult result = DrawListRight(row, col, dataList, rowMerge, colMerge);
                if (rowMerge > 1)
                {
                    result.Row += 1;
                }
                result = DrawBorderByType(row, col, row + rowMerge - 1, result.Col - colMerge - 1, borderType);
                return new CellResult { Row = result.Row, Col = result.Col };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawListRight] row({row}) column({col})");
            }
        }

        public CellResult DrawListRight(int row, int col, List<object> dataList, int rowMerge, int colMerge, BorderType borderType, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                CellResult result = DrawListRight(row, col, dataList, rowMerge, colMerge, borderType);
                return SetAlign(rowStart, colStart, vAlign, hAlign, result.Row, result.Col);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawListRight] row({row}) column({col})");
            }
        }

        //-------------------------------------------------------------- HEADER RIGHT FOR TABLE ----------------------------------------------------------------------

        private CellResult DrawHeaderRightForTable(int row, int col, object obj, int rowMerge, int colMerge, BorderType borderType)
        {
            if (obj != null)
            {
                try
                {
                    rowStart = row;
                    colStart = col;
                    int rowMergeStart = rowMerge;

                    Type type = obj.GetType();
                    PropertyInfo[] properties = type.GetProperties();

                    foreach (PropertyInfo property in properties)
                    {
                        ws.Cells[row, col].Value = property.Name;
                        ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Merge = true;
                        ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Style.WrapText = true;
                        rowMerge = rowMergeStart;
                        col += colMerge;
                    }
                    DrawBorderByType(rowStart, colStart, row + rowMerge, col - 1, borderType);
                    return new CellResult { Row = rowStart + rowMerge, Col = colStart };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawHeaderRightForTable] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawHeaderRightForTable] object/list is NULL. row({row}) column({col})");
            }
        }

        private CellResult DrawHeaderRightForTable(int row, int col, object obj, int rowMerge, int colMerge, BorderType borderType, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                CellResult result = DrawHeaderRightForTable(row, col, obj, rowMerge, colMerge, borderType);
                SetAlign(rowStart, colStart, vAlign, hAlign, result.Row, result.Col);
                return new CellResult { Row = rowStart + rowMerge, Col = colStart };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawHeaderRightForTable] row({row}) column({col})");
            }
        }

        //----------------------------------------------------------------------- TEXT ------------------------------------------------------------------------------

        public CellResult DrawTextStyle(int row, int col, string text, int size, bool isBold, Color color)
        {
            try
            {
                ws.Cells[row, col].Value = text;
                ws.Cells[row, col].Style.Font.Size = size;
                ws.Cells[row, col].Style.Font.Bold = isBold;
                ws.Cells[row, col].Style.Font.Color.SetColor(color);

                return new CellResult { Row = row, Col = col };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawTextStyle] row({row}) column({col})");
            }
        }

        public CellResult DrawTextStyle(int row, int col, string text, int size, bool isBold, Color color, int rowTo, int colTo, BorderType borderType,
            ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                DrawTextStyle(row, col, text, size, isBold, color);
                ws.Cells[row, col, rowTo, colTo].Merge = true;
                ws.Cells[row, col, rowTo, colTo].Style.WrapText = true;

                DrawBorderByType(row, col, rowTo, colTo, borderType);
                SetAlign(row, col, vAlign, hAlign, rowTo, colTo);
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawTextStyle] row({rowTo}) column({colTo})");
            }
        }

        public CellResult DrawText(int row, int col, string text)
        {
            try
            {
                ws.Cells[row, col].Value = text;
                return new CellResult { Row = row, Col = col };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawText] row({row}) column({col})");
            }
        }

        public CellResult DrawText(int row, int col, string text, int rowTo, int colTo)
        {
            try
            {
                ws.Cells[row, col].Value = text;
                ws.Cells[row, col, rowTo, colTo].Merge = true;
                ws.Cells[row, col, rowTo, colTo].Style.WrapText = true;
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawText] row({row}) column({col})");
            }
        }

        public CellResult DrawText(int row, int col, string text, int rowTo, int colTo, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                DrawText(row, col, text, rowTo, colTo);
                SetAlign(row, col, vAlign, hAlign, rowTo, colTo);
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawText] row({rowTo}) column({colTo})");
            }
        }

        public CellResult DrawText(int row, int col, string text, int rowTo, int colTo, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign, BorderType borderType)
        {
            try
            {
                DrawText(row, col, text, rowTo, colTo, vAlign, hAlign);
                return DrawBorderByType(row, col, rowTo, colTo, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawText] row({rowTo}) column({colTo})");
            }
        }

        public CellResult DrawText(int row, int col, string text, int rowTo, int colTo, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign, BorderType borderType, bool isBold)
        {
            try
            {
                DrawText(row, col, text, rowTo, colTo, vAlign, hAlign);
                ws.Cells[row, col].Style.Font.Bold = isBold;
                return DrawBorderByType(row, col, rowTo, colTo, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawText] row({rowTo}) column({colTo})");
            }
        }

        //----------------------------------------------------------------------- DRAW OBJECT ------------------------------------------------------------------------------

        public CellResult DrawObjectStyle(int row, int col, object val, int size, bool isBold, Color color)
        {
            try
            {
                ws.Cells[row, col].Value = val;
                ws.Cells[row, col].Style.Font.Size = size;
                ws.Cells[row, col].Style.Font.Bold = isBold;
                ws.Cells[row, col].Style.Font.Color.SetColor(color);

                return new CellResult { Row = row, Col = col };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawObjectStyle] row({row}) column({col})");
            }
        }

        public CellResult DrawObjectStyle(int row, int col, object val, int size, bool isBold, Color color, int rowTo, int colTo, BorderType borderType,
            ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                DrawObjectStyle(row, col, val, size, isBold, color);
                ws.Cells[row, col, rowTo, colTo].Merge = true;
                ws.Cells[row, col, rowTo, colTo].Style.WrapText = true;

                DrawBorderByType(row, col, rowTo, colTo, borderType);
                SetAlign(row, col, vAlign, hAlign, rowTo, colTo);
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawObjectStyle] row({rowTo}) column({colTo})");
            }
        }

        public CellResult DrawObject(int row, int col, object val)
        {
            try
            {
                ws.Cells[row, col].Value = val;
                return new CellResult { Row = row, Col = col };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawObject] row({row}) column({col})");
            }
        }

        public CellResult DrawObject(int row, int col, object val, int rowTo, int colTo)
        {
            try
            {
                ws.Cells[row, col].Value = val;
                ws.Cells[row, col, rowTo, colTo].Merge = true;
                ws.Cells[row, col, rowTo, colTo].Style.WrapText = true;
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawObject] row({row}) column({col})");
            }
        }

        public CellResult DrawObject(int row, int col, object val, int rowTo, int colTo, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                DrawObject(row, col, val, rowTo, colTo);
                SetAlign(row, col, vAlign, hAlign, rowTo, colTo);
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawObject] row({rowTo}) column({colTo})");
            }
        }

        public CellResult DrawObject(int row, int col, object val, int rowTo, int colTo, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign, BorderType borderType)
        {
            try
            {
                DrawObject(row, col, val, rowTo, colTo, vAlign, hAlign);
                return DrawBorderByType(row, col, rowTo, colTo, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawObject] row({rowTo}) column({colTo})");
            }
        }

        public CellResult DrawObject(int row, int col, object val, int rowTo, int colTo, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign, BorderType borderType, bool isBold)
        {
            try
            {
                DrawObject(row, col, val, rowTo, colTo, vAlign, hAlign);
                ws.Cells[row, col].Style.Font.Bold = isBold;
                return DrawBorderByType(row, col, rowTo, colTo, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawObject] row({rowTo}) column({colTo})");
            }
        }

        //----------------------------------------------------------------------- TABLE DOWN ------------------------------------------------------------------------

        public CellResult DrawTableDown(int row, int col, List<object> dataList)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    colStart = col;

                    foreach (object data in dataList)
                    {
                        Type dataType = data.GetType();
                        PropertyInfo[] properties = dataType.GetProperties();

                        foreach (PropertyInfo property in properties)
                        {
                            object propertyValue = property.GetValue(data);
                            ws.Cells[row, col].Value = propertyValue;
                            col++;
                        }
                        row++;
                        col = colStart;
                    }
                    return new CellResult { Row = row, Col = col };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawTableDown] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawTableDown] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawTableDown(int row, int col, List<object> dataList, bool isWithHeader)
        {
            try
            {
                if (isWithHeader)
                {
                    DrawHeaderRight(row, col, dataList[0]);
                    row++;
                }
                return DrawTableDown(row, col, dataList);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawTableDown] row({row}) column({col})");
            }
        }

        public CellResult DrawTableDownMerge(int row, int col, List<object> dataList, int rowMerge, int colMerge)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    rowStart = row;
                    colStart = col;
                    rowMerge = rowMerge < 1 ? 1 : rowMerge;
                    colMerge = colMerge < 1 ? 1 : colMerge;
                    int colLast = 0;

                    foreach (object data in dataList)
                    {
                        Type dataType = data.GetType();
                        PropertyInfo[] properties = dataType.GetProperties();

                        foreach (PropertyInfo property in properties)
                        {
                            object propertyValue = property.GetValue(data);
                            ws.Cells[row, col].Value = propertyValue;
                            ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Merge = true;
                            ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Style.WrapText = true;
                            col += colMerge;
                        }
                        row += rowMerge;
                        colLast = col > colLast ? col : colLast;
                        col = colStart;
                    }
                    return new CellResult { Row = row, Col = colLast };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawTableDownMerge] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawTableDownMerge] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawTableDownMerge(int row, int col, List<object> dataList, int rowMerge, int colMerge, BorderType borderType)
        {
            try
            {
                CellResult result = DrawTableDownMerge(row, col, dataList, rowMerge, colMerge);
                return DrawBorderByType(rowStart, colStart, result.Row - 1, result.Col - 1, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawTableDownMerge] row({row}) column({col})");
            }
        }

        public CellResult DrawTableDownMerge(int row, int col, List<object> dataList, int rowMerge, int colMerge, BorderType borderType, bool isWithHeader, ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign)
        {
            try
            {
                CellResult result;
                if (isWithHeader)
                {
                    result = DrawHeaderRightForTable(row, col, dataList[0], rowMerge, colMerge, borderType);
                    result = DrawTableDownMerge(result.Row, result.Col, dataList, rowMerge, colMerge);
                }
                else
                {
                    result = DrawTableDownMerge(row, col, dataList, rowMerge, colMerge);
                }
                SetAlign(row, col, vAlign, hAlign, result.Row, result.Col);
                return DrawBorderByType(rowStart, colStart, result.Row - 1, result.Col - 1, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawTableDownMerge] row({row}) column({col})");
            }
        }

        public CellResult DrawTableDownMergeColor(int row, int col, List<object> dataList, int rowMerge, int colMerge, Color color, BgColorType bgColorType)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    rowStart = row;
                    colStart = col;
                    rowMerge = rowMerge < 1 ? 1 : rowMerge;
                    colMerge = colMerge < 1 ? 1 : colMerge;
                    int colLast = 0;
                    int i = 0;

                    foreach (object data in dataList)
                    {
                        Type dataType = data.GetType();
                        PropertyInfo[] properties = dataType.GetProperties();

                        foreach (PropertyInfo property in properties)
                        {
                            object propertyValue = property.GetValue(data);
                            ws.Cells[row, col].Value = propertyValue;
                            ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Merge = true;
                            ws.Cells[row, col, row + rowMerge - 1, col + colMerge - 1].Style.WrapText = true;
                            col += colMerge;
                        }
                        row += rowMerge;
                        colLast = col > colLast ? col : colLast;
                        col = colStart;

                        if (bgColorType.Equals(BgColorType.Odd))
                        {
                            if (i == 0)
                            {
                                SetBackgroundColor(rowStart, colStart, color, row - 1, colLast - 1);
                            }
                            else if (i % 2 != 0)
                            {
                                SetBackgroundColor(row, col, color, row, colLast - 1);
                            }
                        }

                        if (bgColorType.Equals(BgColorType.Even) && (i % 2 == 0) && (i < dataList.Count - 1))
                        {
                            SetBackgroundColor(row, col, color, row, colLast - 1);
                        }

                        i++;
                    }

                    if (bgColorType == BgColorType.All)
                    {
                        SetBackgroundColor(rowStart, colStart, color, row - 1, colLast - 1);
                    }

                    return new CellResult { Row = row, Col = colLast };
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawTableDownMergeColor] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawTableDownMergeColor] object/list is NULL. row({row}) column({col})");
            }
        }

        public CellResult DrawTableDownMergeColor(int row, int col, List<object> dataList, int rowMerge, int colMerge, BorderType borderType,
            ExcelVerticalAlignment vAlign, ExcelHorizontalAlignment hAlign, Color color, BgColorType bgColorType)
        {
            if (dataList.Count > 0)
            {
                try
                {
                    CellResult result;
                    result = DrawTableDownMergeColor(row, col, dataList, rowMerge, colMerge, color, bgColorType);

                    SetAlign(row, col, vAlign, hAlign, result.Row, result.Col);
                    return DrawBorderByType(rowStart, colStart, result.Row - 1, result.Col - 1, borderType);
                }
                catch (Exception e)
                {
                    throw new Exception($"{e.Message} \n[DrawTableDownMergeColor] row({row}) column({col})");
                }
            }
            else
            {
                throw new Exception($"[DrawTableDownMergeColor] object/list is NULL. row({row}) column({col})");
            }
        }

        //----------------------------------------------------------------------- BACKGROUND COLOR ------------------------------------------------------------------------

        public CellResult SetBackgroundColor(int row, int col, Color color)
        {
            try
            {
                ws.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(color);
                return new CellResult { Row = row, Col = col };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[SetBackgroundColor] row({row}) column({col})");
            }
        }

        public CellResult SetBackgroundColor(int row, int col, Color color, int rowTo, int colTo)
        {
            try
            {
                ws.Cells[row, col, rowTo, colTo].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, col, rowTo, colTo].Style.Fill.BackgroundColor.SetColor(color);
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[SetBackgroundColor] row({row}) column({col})");
            }
        }

        //--------------------------------------------------------------------------- MATH ---------------------------------------------------------------------------

        public double SumLineInc(int row, int col, int rowTo, int colTo)
        {
            double sum = 0;
            foreach (var cell in ws.Cells[row, col, rowTo, colTo])
            {
                double cellVal = 0;
                if (double.TryParse(cell.Value?.ToString(), out cellVal))
                {
                    sum += cellVal;
                }
            }
            return sum;
        }

        public double SumLineDec(int row, int col, int rowTo, int colTo)
        {
            double sum = 0;
            foreach (var cell in ws.Cells[row, col, rowTo, colTo])
            {
                double cellVal = 0;
                if (double.TryParse(cell.Value?.ToString(), out cellVal))
                {
                    sum -= cellVal;
                }
            }
            return sum;
        }

        //--------------------------------------------------------------------------- GET VALUE ---------------------------------------------------------------------------

        public object GetValue(int row, int col)
        {
            return ws.GetValue(row, col);
        }

        public string GetValueString(int row, int col)
        {
            return ws.GetValue(row, col).ToString();
        }

        public int GetValueInt(int row, int col)
        {
            return Convert.ToInt32(ws.GetValue(row, col));
        }

        public Int16 GetValueInt16(int row, int col)
        {
            return Convert.ToInt16(ws.GetValue(row, col));
        }

        public Int32 GetValueInt32(int row, int col)
        {
            return Convert.ToInt32(ws.GetValue(row, col));
        }

        public Int64 GetValueInt64(int row, int col)
        {
            return Convert.ToInt64(ws.GetValue(row, col));
        }

        public long GetValueLong(int row, int col)
        {
            return Convert.ToInt64(ws.GetValue(row, col));
        }

        public decimal GetValueDecimal(int row, int col)
        {
            return Convert.ToDecimal(ws.GetValue(row, col));
        }

        public float GetValueFloat(int row, int col)
        {
            return Convert.ToInt64(ws.GetValue(row, col));
        }

        public double GetValueDouble(int row, int col)
        {
            return Convert.ToDouble(ws.GetValue(row, col));
        }

        public bool GetValueBool(int row, int col)
        {
            return Convert.ToBoolean(ws.GetValue(row, col));
        }

        public byte GetValueByte(int row, int col)
        {
            return Convert.ToByte(ws.GetValue(row, col));
        }

        //----------------------------------------------------------------------- BORDER TYPE ------------------------------------------------------------------------

        public CellResult DrawBorder(int rowStart, int colStart, int rowTo, int colTo, BorderType borderType)
        {
            try
            {
                return DrawBorderByType(rowStart, colStart, rowTo, colTo, borderType);
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawBorder] row({rowStart}) column({colStart})");
            }
        }

        public CellResult DrawBorderByType(int rowStart, int colStart, int rowTo, int colTo, BorderType borderType)
        {
            try
            {
                switch (borderType)
                {
                    case BorderType.NoBorderAll:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Top.Style = ExcelBorderStyle.None;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Right.Style = ExcelBorderStyle.None;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.None;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Left.Style = ExcelBorderStyle.None;
                        }
                        break;
                    case BorderType.NoBorderAround:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.BorderAround(ExcelBorderStyle.None);
                        }
                        break;
                    case BorderType.BorderAroundThin:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                        break;
                    case BorderType.BorderAroundThick:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                        }
                        break;
                    case BorderType.BorderAllThin:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        }
                        break;
                    case BorderType.BorderAllThick:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                        }
                        break;
                    case BorderType.BorderAroundDotted:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.BorderAround(ExcelBorderStyle.Dotted);
                        }
                        break;
                    case BorderType.BorderAroundDashed:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.BorderAround(ExcelBorderStyle.Dashed);
                        }
                        break;
                    case BorderType.BorderAllDotted:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Right.Style = ExcelBorderStyle.Dotted;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Dotted;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Left.Style = ExcelBorderStyle.Dotted;
                        }
                        break;
                    case BorderType.BorderAllDashed:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Top.Style = ExcelBorderStyle.Dashed;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Right.Style = ExcelBorderStyle.Dashed;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Dashed;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Left.Style = ExcelBorderStyle.Dashed;
                        }
                        break;
                    case BorderType.BorderBottomThin:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        }
                        break;
                    case BorderType.BorderBottomThick:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                        }
                        break;
                    case BorderType.BorderBottomDotted:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Dotted;
                        }
                        break;
                    case BorderType.BorderBottomDashed:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Dashed;
                        }
                        break;
                    case BorderType.BorderBottomDouble:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                        }
                        break;
                    case BorderType.BorderAroundDouble:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        }
                        break;
                    case BorderType.BorderAllDouble:
                        {
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Top.Style = ExcelBorderStyle.Double;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Right.Style = ExcelBorderStyle.Double;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                            ws.Cells[rowStart, colStart, rowTo, colTo].Style.Border.Left.Style = ExcelBorderStyle.Double;
                        }
                        break;
                }
                return new CellResult { Row = rowTo, Col = colTo };
            }
            catch (Exception e)
            {
                throw new Exception($"{e.Message} \n[DrawBorderByType] row({rowStart}) column({colStart})");
            }
        }

        //eof
    }
}