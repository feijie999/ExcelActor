using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelActor.Interface;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style;
using Orleans.Concurrency;

namespace ExcelActor.Grain
{
    //[StatelessWorker]
    public class ExcelGrain : Orleans.Grain<object>,IExcelGrain
    {
        private ExcelPackage _excelPackage;
        private object _jsonTemplate;

        public override async Task OnActivateAsync()
        {
            _excelPackage = new ExcelPackage();
           // await ReadStateAsync();
            var bytes = State as byte[];
            if (bytes != null && bytes.Length > 0)
            {
                using (var stream = new MemoryStream(bytes))
                {
                    _excelPackage.Load(stream);
                }
            }

            await base.OnActivateAsync();
        }

        public Task Load(byte[] excelBytes)
        {
            State = excelBytes;
            using (var stream = new MemoryStream(excelBytes))
            {
                _excelPackage.Load(stream);
            }
            return Task.CompletedTask;
        }

        [AlwaysInterleave]
        public Task<string> ExportAllToText()
        {
            var obj = ExportAllToJson();
            return Task.FromResult(JsonConvert.SerializeObject(obj));
        }

        public object ExportAllToJson()
        {
            var tables = new List<object>();
            for (var i = 1; i <= _excelPackage.Workbook.Worksheets.Count; i++)
            {
                var table = ExportToJson(i);
                tables.Add(table);
            }

            _jsonTemplate = tables;

            return tables;
        }

        public object ExportToJson(int sheetIndex = 1)
        {
            var sheet = _excelPackage.Workbook.Worksheets[sheetIndex];
            var cells = sheet.Cells;


            // 填值后计算所有工作表
            //_excelPackage.Workbook.CalcMode = ExcelCalcMode.Manual;
            //cells["AH9"].Value = 100;
            //cells["AH10"].Value = 300;

            //cells["AN9"].Value = 100;
            //cells["AN10"].Value = 300;
            //_excelPackage.Workbook.Calculate();

            var rows = new List<Dictionary<string, object>>();
            for (var i = sheet.Dimension.Start.Row; i <= sheet.Dimension.End.Row; i++)
            {
                var row = new Dictionary<string, object>();
                var excelRow = sheet.Row(i);
                var rowStyle = new Dictionary<string, object>();
                rowStyle.AddNotNull("height", excelRow.Hidden ? "0" : excelRow.Height.PixelHeight());

                var cols = new List<Dictionary<string, object>>();

                row.AddNotNull("row", i)
                    .AddNotNull("hidden", excelRow.Hidden)
                    .AddNotNullOrSpace("style", rowStyle)
                    .AddNotNull("cols", cols);
                for (var j = sheet.Dimension.Start.Column; j <= sheet.Dimension.End.Column; j++)
                {
                    var excelColumn = sheet.Column(j);
                    var cell = cells[i, j];

                    var col = new Dictionary<string, object>();

                    col.AddNotNull("row", i).AddNotNull("col", j)
                        .AddNotNullOrSpace("formula", cell.Formula);
                    if (!cell.IsRichText)
                    {
                        // 处理输入标志
                        if (cell.Value?.ToString() == "{N}")
                        {
                            col.AddNotNull("isInput", true)
                                .AddNotNull("type", "number");
                        }
                        else if (cell.Value?.ToString() == "{S}")
                        {
                            col.AddNotNull("isInput", true)
                                .AddNotNull("type", "string");

                        }
                        else if (cell.Value is KeyValuePair<bool, string>)
                        {
                            cell.Value = ((KeyValuePair<bool, string>)cell.Value).Value;
                            col.AddNotNull("value", cell.Value);
                            col.AddNotNull("isInput", true)
                                .AddNotNull("type", "string");
                        }
                        else
                            col.AddNotNull("value", cell.Value);
                    }
                    else
                    {
                        var richText = new List<Dictionary<string, object>>();
                        foreach (var text in cell.RichText)
                        {
                            var theRichText = new Dictionary<string, object>();

                            var style = new Dictionary<string, object>();
                            style.AddNotNullOrSpace("fontSize", text.Size.PixelSize())
                                .AddNotNullOrSpace("fontFamily", $"'{text.FontName}'");

                            theRichText.AddNotNullOrSpace("text", text.Text)
                                .AddNotNull("style", style);

                            richText.Add(theRichText);
                        }
                        col.AddNotNull("isRichText", cell.IsRichText)
                            .AddNotNull("value", richText);
                    }
                    
                    var colStyle = new Dictionary<string, object>();
                    var width = excelColumn.Width.PixelWidth(_excelPackage.Workbook.MaxFontWidth);
                    var backgroundColor = CreateColor(cell.Style.Fill.BackgroundColor);
                    var fontSize = cell.Style.Font.Size.PixelSize();
                    var fontFamily = cell.Style.Font.Name;
                    var fontWeight = cell.Style.Font.Bold ? "bold" : "";
                    colStyle
                        .AddNotNull("width", width)
                        .AddNotNull("backgroundColor", backgroundColor)
                        .AddNotNull("fontSize", fontSize)
                        .AddNotNullOrSpace("fontFamily", fontFamily)
                        .AddNotNullOrSpace("fontWeight", fontWeight);

                    switch (cell.Style.HorizontalAlignment)
                    {
                        case ExcelHorizontalAlignment.Distributed:
                            colStyle.AddNotNull("textAlign", "justify")
                                .AddNotNull("textAlignLast", "justify");
                            break;
                        default:
                            colStyle.AddNotNull("textAlign", cell.Style.HorizontalAlignment.ToString());
                            break;
                    }

                    if (cell.Merge)
                    {
                        var merge = GetMergeRange(sheet, i, j);

                        if (merge != null && i == merge.Item1 && j == merge.Item2)
                        {
                            col.AddNotNull("hidden", excelRow.Hidden || false)
                                .AddNotNull("rowspan", merge.Item3)
                                .AddNotNull("colspan", merge.Item4);
                        }
                        else
                        {
                            col.AddNotNull("hidden", true);
                        }
                    }
                    else
                    {
                        col.AddNotNull("hidden", excelRow.Hidden || false);
                    }

                    col.AddNotNullOrSpace("style", colStyle);
                    cols.Add(col);
                }
                rows.Add(row);
            }
           
            rows.Add(AppendRow(sheet));

            var table = new Dictionary<string, object>
            {
                { "name", sheet.Name },
                { "rows", rows }
            };

            AddBorder(table, sheet);

            AddPicture(table, sheet);
            table = Clean(table, sheet);

            AdjustBorder(table, sheet);

            return table;
        }

        private void AddPicture(Dictionary<string, object> table, ExcelWorksheet sheet)
        {
            foreach (ExcelPicture drawing in sheet.Drawings)
            {
                var cell = GetTableCell(table, drawing.From.Row + 1, drawing.From.Column + 1);
                using (var ms = new MemoryStream())
                {
                    drawing.Image.Save(ms, drawing.Image.RawFormat);
                    var buffer = new byte[ms.Length];
                    ms.Position = 0;
                    ms.Read(buffer, 0, buffer.Length);

                    var picture = new Dictionary<string, object>
                    {
                        { "format", drawing.ImageFormat.ToString() },
                        { "image", Convert.ToBase64String(buffer) }
                    };
                    cell.Add("picture", picture);
                    cell["hidden"] = false;
                }
            }
        }

        // 生成边框线
        private void AddBorder(Dictionary<string, object> table, ExcelWorksheet sheet)
        {
            foreach (var row in table["rows"] as List<Dictionary<string, object>>)
            {
                foreach (var col in row["cols"] as List<Dictionary<string, object>>)
                {
                    int i = (int)col["row"],
                        j = (int)col["col"];

                    int rs = col.ContainsKey("rowspan") ? (int)col["rowspan"] : 1;
                    int cs = col.ContainsKey("colspan") ? (int)col["colspan"] : 1;

                    var cell = sheet.Cells[i, j];
                    var style = col["style"] as Dictionary<string, object>;
                    // 左边框、顶边框
                    style.AddNotNullOrSpace("borderLeft", CreateBorder(cell.Style.Border.Left))
                        .AddNotNullOrSpace("borderTop", CreateBorder(cell.Style.Border.Top));

                    // 右边框、底边框
                    cell = sheet.Cells[i, j + cs - 1];
                    style.UpdateNotNullOrSpace("borderRight", CreateBorder(cell.Style.Border.Right));
                    cell = sheet.Cells[i + rs - 1, j];
                    style.UpdateNotNullOrSpace("borderBottom", CreateBorder(cell.Style.Border.Bottom));
                }
            }
        }

        // 调整边框线
        private void AdjustBorder(Dictionary<string, object> table, ExcelWorksheet sheet)
        {
            foreach (var row in table["rows"] as List<Dictionary<string, object>>)
            {
                foreach (var col in row["cols"] as List<Dictionary<string, object>>)
                {
                    int i = (int)col["row"],
                        j = (int)col["col"];

                    int rs = col.ContainsKey("rowspan") ? (int)col["rowspan"] : 1;
                    int cs = col.ContainsKey("colspan") ? (int)col["colspan"] : 1;

                    var cell = sheet.Cells[i, j];
                    var style = col["style"] as Dictionary<string, object>;

                    for (var k = i; k < i + rs; k++)
                    {
                        var left = GetMergeRange(sheet, k, j - 1);
                        if (left != null)
                        {
                            var target = GetTableCell(table, left.Item1, left.Item2);
                            if (target != null)
                            {
                                var leftStyle = target["style"] as Dictionary<string, object>;

                                if (!style.ContainsKey("borderLeft") && leftStyle.ContainsKey("borderRight"))
                                {
                                    if (rs == left.Item3)
                                    {
                                        style.AddNotNullOrSpace("borderLeft", leftStyle["borderRight"] as string);
                                    }
                                }
                                else if (style.ContainsKey("borderLeft") && !leftStyle.ContainsKey("borderRight"))
                                {
                                    leftStyle.AddNotNullOrSpace("borderRight", style["borderLeft"] as string);
                                }

                                if (style.ContainsKey("borderLeft") && leftStyle.ContainsKey("borderRight"))
                                {
                                    leftStyle.Remove("borderRight");
                                }
                            }
                        }
                    }

                    bool removeTop = false;
                    for (var k = j; k < j + cs; k++)
                    {
                        var top = GetMergeRange(sheet, i - 1, k);

                        if (top != null)
                        {
                            var target = GetTableCell(table, top.Item1, top.Item2);
                            if (target != null)
                            {
                                var topStyle = target["style"] as Dictionary<string, object>;

                                if (!style.ContainsKey("borderTop") && topStyle.ContainsKey("borderBottom"))
                                {
                                    if (cs == top.Item4)
                                    {
                                        style.AddNotNullOrSpace("borderTop", topStyle["borderBottom"] as string);
                                    }
                                }
                                else if (style.ContainsKey("borderTop") && !topStyle.ContainsKey("borderBottom"))
                                {
                                    if (cs >= top.Item4)
                                    {
                                        topStyle.AddNotNullOrSpace("borderBottom", style["borderTop"] as string);
                                    }
                                }

                                if (topStyle.ContainsKey("borderBottom"))
                                    removeTop = true;
                            }
                        }
                    }

                    if (removeTop)
                        style.Remove("borderTop");
                }
            }
        }

        // 清理多余的属性
        private Dictionary<string, object> Clean(Dictionary<string, object> table, ExcelWorksheet sheet)
        {
            foreach (var row in table["rows"] as List<Dictionary<string, object>>)
            {
                row["cols"] = (row["cols"] as List<Dictionary<string, object>>)
                    .Where(p => (p.ContainsKey("hidden") && p["hidden"] as bool? == false))
                    .Select(p => p.Delete("hidden").Delete("colspan", 1).Delete("rowspan", 1)).ToList();
            }

            return table;
        }

        private Dictionary<string, object> GetTableCell(Dictionary<string, object> table, int rowIndex, int colIndex)
        {
            foreach (var row in table["rows"] as List<Dictionary<string, object>>)
            {
                if ((int)row["row"] != rowIndex)
                    continue;

                foreach (var col in row["cols"] as List<Dictionary<string, object>>)
                {
                    if ((int)col["col"] != colIndex)
                        continue;
                    return col;
                }
            }

            return null;
        }

        private string CreateColor(ExcelColor color)
        {
            var rgb = color.LookupColor(color);
            if (!string.IsNullOrWhiteSpace(color.Rgb))
            {
                return $"#{color.Rgb.Substring(2)}";
            }
            else if (color.Indexed > 0 && color.Indexed < 64)
            {

                return $"#{rgb.Substring(3)}";
            }

            return null;
        }

        private string CreateBorder(ExcelBorderItem item)
        {
            switch (item.Style)
            {
                case ExcelBorderStyle.None:
                    return "";
                case ExcelBorderStyle.Thin:
                    return "1px solid";
                case ExcelBorderStyle.Medium:
                    return "2px solid";
            }

            return null;
        }
        /// <summary>
        /// 追加一行，hack
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, object> AppendRow(ExcelWorksheet sheet)
        {
            var cols = new List<Dictionary<string, object>>();
            for (var j = sheet.Dimension.Start.Column; j <= sheet.Dimension.End.Column; j++)
            {
                var cell = sheet.Cells[1, j];
                var col = new Dictionary<string, object>();
                var colStyle = new Dictionary<string, object>();
                var width = sheet.Column(j).Width.PixelWidth(_excelPackage.Workbook.MaxFontWidth);
                colStyle.AddNotNull("width", width);
                col.AddNotNull("hidden", false)
                    .AddNotNull("row", sheet.Dimension.End.Row + 1)
                    .AddNotNull("col", j)
                    .AddNotNull("style", colStyle);

                cols.Add(col);
            }
            var row = new Dictionary<string, object>();
            row.AddNotNull("row", sheet.Dimension.End.Row + 1)
                .AddNotNull("style", new Dictionary<string, object> { { "height", 0 } })
                .AddNotNull("cols", cols);

            return row;
        }
        public ExcelPackage ImportFromJson(string json)
        {
            return null;
        }

        private Tuple<int, int, int, int> GetMergeRange(ExcelWorksheet sheet, int row, int col)
        {
            if (row < 1 || col < 1)
            {
                return null;
            }

            if (!sheet.Cells[row, col].Merge)
            {
                return new Tuple<int, int, int, int>(row, col, 1, 1);
            }

            foreach (var cell in sheet.MergedCells)
            {
                var merge = GetRange(cell);
                if (row >= merge.Item1 && row <= merge.Item3 && col >= merge.Item2 && col <= merge.Item4)
                    return new Tuple<int, int, int, int>(merge.Item1, merge.Item2, merge.Item3 - merge.Item1 + 1, merge.Item4 - merge.Item2 + 1);
            }

            return null;
        }

        /// <summary>
        /// 返回合并单元
        /// </summary>
        /// <param name="rang">合并单元范围串</param>
        /// <returns></returns>
        private Tuple<int, int, int, int> GetRange(string rang)
        {
            var tmp = rang.Split(":");
            int startRow = RowIndex(tmp[0]);
            int startColumn = ColumnIndex(ColumnName(tmp[0]));
            int endRow = RowIndex(tmp[1]);
            int endColumn = ColumnIndex(ColumnName(tmp[1]));

            return new Tuple<int, int, int, int>(startRow, startColumn, endRow, endColumn);
        }

        private string ColumnName(string cellName)
        {
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellName);
            return match.Value;
        }

        private int ColumnIndex(string columnName)
        {
            int index = 0;
            foreach (var ch in columnName.ToUpper())
            {
                index *= 26;
                index += (ch - 'A') + 1;
            }

            return index;
        }

        private int RowIndex(string cellName)
        {
            var regex = new Regex(@"\d+");
            var match = regex.Match(cellName);
            return int.Parse(match.Value);
        }

        /// <summary>
        /// 从填充了数据的Json表格提取数据单元
        /// 返回的数据格式 [{row:x,col:y,value: 'data'}....]
        /// </summary>
        /// <param name="sourceTables"></param>
        /// <returns></returns>
        public object GetDataWithChanged(JArray tablesWithData)
        {
            var resultData = new JArray();

            for (var tableIndex = 0; tableIndex < tablesWithData.Count; tableIndex++)
            {
                var rows = (JArray)tablesWithData[tableIndex]["rows"];
                var templateRows = (JArray)((JArray)_jsonTemplate)[tableIndex]["rows"];
                var dataRows = new JArray();

                for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    var cols = (JArray)rows[rowIndex]["cols"];
                    var templateCols = (JArray)templateRows[rowIndex]["cols"];
                    var dataCols = new JArray();

                    for (var colIndex = 0; colIndex < cols.Count; colIndex++)
                    {
                        var col = cols[colIndex];
                        var templateCol = templateCols[colIndex];

                        if (col["value"] == null) continue;
                        if (templateCol["value"] != null && col["value"].ToString() == templateCol["value"].ToString()) continue;

                        var dataCol = new JObject();
                        dataCol.Add("row", col["row"]);
                        dataCol.Add("col", col["col"]);
                        dataCol.Add("value", col["value"]);

                        dataCols.Add(dataCol);
                        Console.WriteLine($"({col["row"]},{col["col"]}) = {col["value"]} ==> {templateCol["value"]}");
                    }
                    if (dataCols.Count > 0)
                    {
                        dataRows.Add(new JObject { { "row", rowIndex }, { "cols", dataCols } });
                    }
                }
                if (dataRows.Count > 0)
                {
                    resultData.Add(new JObject { { "rows", dataRows } });
                }
            }

            return resultData;
        }

        /// <summary>
        /// 填充数据
        /// 数据格式 [{row:x,col:y,value: 'data'}....]
        /// </summary>
        /// <param name="data"></param>
        public object FillingData(JArray data)
        {
            for (var tableIndex = 0; tableIndex < data.Count; tableIndex++)
            {
                var rows = (JArray)data[tableIndex]["rows"];
                var cells = _excelPackage.Workbook.Worksheets[tableIndex + 1].Cells;

                for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    var cols = (JArray)rows[rowIndex]["cols"];
                    for (var colIndex = 0; colIndex < cols.Count; colIndex++)
                    {
                        var col = cols[colIndex];
                        var r = int.Parse(col["row"].ToString());
                        var c = int.Parse(col["col"].ToString());
                        cells[r, c].Value = new KeyValuePair<bool, string>(true, col["value"].ToString());
                    }
                }
            }

            _excelPackage.Workbook.CalcMode = ExcelCalcMode.Manual;
            _excelPackage.Workbook.Calculate();

            return ExportAllToJson();
        }

        /// <summary>
        /// 计算公式
        /// </summary>
        public void RecalcFormula()
        {
            _excelPackage.Workbook.Calculate();
        }

        public async Task<string> Test(string name)
        {
            await Task.Delay(3000);
            Console.WriteLine(name);
            return "Hi " + name;
        }

        public override async Task OnDeactivateAsync()
        {
            _excelPackage.Dispose();
            await WriteStateAsync();
            await base.OnDeactivateAsync();
        }
    }
}
