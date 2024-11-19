# MagerExcel

MagerExcel is a C# library for working with Excel files.

## Installation

Install the NuGet package:

```bash
dotnet add package MagerExcel

Here's how to use MyLibrary (just make an ajax request to this endpoint):

using OfficeOpenXml;
using OfficeOpenXml.Style;
using SAMPLE_APP.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Mvc;
using System.Drawing;
using MagerExcel;
using System.Linq;
using System.Data.Linq;
using OfficeOpenXml.Drawing.Chart;

namespace SAMPLE_APP.Controllers
{
    public class MagerExcelController : Controller
    {
        public ActionResult ExportTest3()
        {
            var memoryStream = new MemoryStream();
            using (var excelPackage = new ExcelPackage(memoryStream))
            {
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add("Sheet1");
                clsMagerExcel mae = new clsMagerExcel(ws);
                //clsMagerExcel.CellResult cell;

                //--
                ws.Cells["A1"].Value = "Category";
                ws.Cells["B1"].Value = "Value";
                ws.Cells["A2"].Value = "A";
                ws.Cells["B2"].Value = 10;
                ws.Cells["A3"].Value = "B";
                ws.Cells["B3"].Value = 20;

                var chart = ws.Drawings.AddChart("BarChart", eChartType.ColumnClustered);
                chart.SetPosition(1, 0, 2, 0);
                chart.SetSize(400, 300);
                var series1 = chart.Series.Add(ws.Cells["B2:B3"], ws.Cells["A2:A3"]);
                series1.Header = "OK Product";
                //--
                // Menambahkan data contoh untuk series kedua
                //var dataSeries2 = new List<int> { 30, 40, 50, 60, 70 };

                //// Atur sumbu X dan Y untuk series kedua
                //var series2 = chart.Series.Add(worksheet.Cells["C1:C5"], worksheet.Cells["A1:A5"]);
                //series2.Header = "Good Product";


                Session["asd"] = excelPackage.GetAsByteArray();
            }

            if (Session["asd"] != null)
            {
                byte[] data = Session["asd"] as byte[];
                return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "asd_" + Convert.ToDateTime(DateTime.Now).ToString("yyyyMMdd") + ".xlsx");
            }
            else
            {
                return new EmptyResult();
            }
        }

        public ActionResult ExportTest2()
        {
            var memoryStream = new MemoryStream();
            using (var excelPackage = new ExcelPackage(memoryStream))
            {
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add("Sheet1");
                clsMagerExcel mae = new clsMagerExcel(ws);
                //clsMagerExcel.CellResult cell;

                for (int i = 1; i <= 10; i++)
                {
                    ws.Cells[1, i].Value = i;
                    ws.Cells[2, i].Value = i+20;
                }

                OfficeOpenXml.Table.PivotTable.ExcelPivotTable pvt = ws.PivotTables.Add(ws.Cells["A1:I1"], ws.Cells["A2:I2"], "Pvt01");

                for (int i = 1; i < 8; i++)
                {
                    pvt.ColumnFields[i].Name = $"Col{i}";
                }
                //var chartBar = ws.Drawings.AddChart("chartBar", OfficeOpenXml.Drawing.Chart.eChartType.BarClustered, pvt);
                var chartDonat = ws.Drawings.AddChart("chartDonat", OfficeOpenXml.Drawing.Chart.eChartType.Doughnut, pvt);
                chartDonat.SetPosition(10, 0, 3, 0);
                chartDonat.SetSize(400, 300);
                chartDonat.Series.Add(ws.Cells["B22:B23"], ws.Cells["A22:A23"]);
                chartDonat.Title.Text = "Chart Donat";


                Session["asd"] = excelPackage.GetAsByteArray();
            }

            if (Session["asd"] != null)
            {
                byte[] data = Session["asd"] as byte[];
                return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "asd_" + Convert.ToDateTime(DateTime.Now).ToString("yyyyMMdd") + ".xlsx");
            }
            else
            {
                return new EmptyResult();
            }
        }

        public ActionResult ExportTest1()
        {
            var memoryStream = new MemoryStream();
            using (var excelPackage = new ExcelPackage(memoryStream))
            {
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add("Sheet1");
                clsMagerExcel mae = new clsMagerExcel(ws);
                //clsMagerExcel.CellResult cell;

                using (var db = new DataContext(clsDBHelper.DBCS))
                {
                    var tbl = new List<object>();
                    foreach (var obj in db.ExecuteQuery<clsProcessTablePackingReport.ShrinkHeater>("sp_ShrinkHeater_Test1").ToList())
                    {
                        tbl.Add(obj);
                    }
                    mae.DrawTableRightMerge(1, 2, tbl, 2, 3, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center, clsMagerExcel.BorderType.BorderAllThin);
                }
                Session["asd"] = excelPackage.GetAsByteArray();
            }

            if (Session["asd"] != null)
            {
                byte[] data = Session["asd"] as byte[];
                return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "asd_" + Convert.ToDateTime(DateTime.Now).ToString("yyyyMMdd") + ".xlsx");
            }
            else
            {
                return new EmptyResult();
            }
        }

        public ActionResult ExportDownload()
        {
            try
            {
                var memoryStream = new MemoryStream();
                using (var excelPackage = new ExcelPackage(memoryStream))
                {
                    ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add("Sheet1");
                    clsMagerExcel mae = new clsMagerExcel(ws);
                    clsMagerExcel.CellResult cell;

                    var MasalahTrouble = new List<object> 
                    {
                        new clsProcessTablePackingReport.MasalahTrouble
                        {
                            Action = "act1",
                            Item = "item1",
                            Menit = 60,
                            Start = "07.00",
                            Stop = "19.00"
                        },
                        new clsProcessTablePackingReport.MasalahTrouble
                        {
                            Action = "act2",
                            Item = "item2",
                            Menit = 62,
                            Start = "02.00",
                            Stop = "15.00"
                        },
                        new clsProcessTablePackingReport.MasalahTrouble
                        {
                            Action = "act5",
                            Item = "item5",
                            Menit = 65,
                            Start = "05.00",
                            Stop = "10.00"
                        }
                    };

                    mae.DrawText(1, 1, "wow12345");
                    
                    mae.DrawTableDown(21, 1, MasalahTrouble);

                    mae.DrawTableDown(31, 1, MasalahTrouble, true);

                    mae.DrawHeaderDown(41, 1, MasalahTrouble[0]);

                    mae.DrawHeaderRight(51, 1, MasalahTrouble[0], 2, 5);

                    mae.DrawHeaderRight(61, 1, MasalahTrouble[0], 2, 5, clsMagerExcel.BorderType.BorderAllDashed);

                    mae.DrawTableDownMerge(71, 1, MasalahTrouble, 2, 3);

                    mae.DrawTableDownMerge(81, 1, MasalahTrouble, 2, 3, clsMagerExcel.BorderType.BorderAllDotted);

                    cell = mae.DrawTableDownMerge(91, 1, MasalahTrouble, 2, 3, clsMagerExcel.BorderType.BorderAroundDashed, true, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    cell = mae.DrawTableDownMerge(cell.Row+5, 1, MasalahTrouble, 2, 3, clsMagerExcel.BorderType.BorderAllDotted, true, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    cell = mae.DrawText(cell.Row + 5, 1, "INI ADALAH MAGER EXCEL", cell.Row + 5, 10, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    cell = mae.DrawText(cell.Row + 5, 1, "INI ADALAH MAGER EXCEL 123 HORE", cell.Row + 5, 10, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center, clsMagerExcel.BorderType.BorderAroundThick);

                    cell = mae.DrawHeaderRight(cell.Row + 5, 3, MasalahTrouble[0], 2, 5);

                    cell = mae.DrawHeaderRight(cell.Row + 5, 3, MasalahTrouble[0], 1, 5);

                    cell = mae.DrawHeaderRight(cell.Row + 5, 3, MasalahTrouble[0], 1, 1);

                    cell = mae.DrawHeaderRight(cell.Row + 5, 2, MasalahTrouble[0], 2, 5, clsMagerExcel.BorderType.BorderAllDashed, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    cell = mae.DrawHeaderRight(cell.Row + 5, 2, MasalahTrouble[0], 2, 5, clsMagerExcel.BorderType.BorderAllDotted, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Left);

                    cell = mae.DrawHeaderRight(cell.Row + 5, 2, MasalahTrouble[0], 2, 5, clsMagerExcel.BorderType.BorderAllThin, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Right);

                    cell = mae.DrawHeaderRight(cell.Row + 5, 2, MasalahTrouble[0], 1, 3, clsMagerExcel.BorderType.BorderAllThick, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    cell = mae.DrawTableDownMerge(cell.Row + 5, 3, MasalahTrouble, 3, 6, clsMagerExcel.BorderType.BorderAllThick, true, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    cell = mae.DrawListDown(cell.Row + 5, 2, new List<object> { "item 1 ", "item 2", "aloha", "wow 123" });

                    cell = mae.DrawListDown(cell.Row + 5, 2, new List<object> { "item 1 ", "item 2", "aloha", "wow 123" }, 3, 8);

                    cell = mae.DrawListDown(cell.Row + 5, 2, new List<object> { "item 1 ", "item 2", "aloha", "wow 123" }, 2, 4, clsMagerExcel.BorderType.BorderAroundThin);

                    cell = mae.DrawListDown(cell.Row + 5, 2, new List<object> { "item 1 ", "item 2", "aloha", "wow 123" }, 2, 4, clsMagerExcel.BorderType.BorderAllThin, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    cell.Row += 3;
                    cell = mae.DrawTableDownMergeColor(cell.Row, 3, MasalahTrouble, 3, 5, clsMagerExcel.BorderType.BorderAllThin, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Right, Color.MediumAquamarine, clsMagerExcel.BgColorType.Even);
                    
                    cell.Row += 3;
                    cell = mae.DrawTableDownMergeColor(cell.Row, 3, MasalahTrouble, 3, 5, clsMagerExcel.BorderType.BorderAllThin, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Right, Color.MediumPurple, clsMagerExcel.BgColorType.Odd);

                    cell.Row += 3;
                    cell = mae.DrawTableDownMergeColor(cell.Row, 3, MasalahTrouble, 3, 5, clsMagerExcel.BorderType.BorderAllThin, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Right, Color.GreenYellow, clsMagerExcel.BgColorType.All);

                    cell.Row += 3;

                    cell = mae.DrawListRight(cell.Row, 1, new List<object>
                    {
                        "Masalah",
                        "Action",
                        "Stop",
                        "Start"
                    }, 1, 3, clsMagerExcel.BorderType.BorderAllThin, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    cell.Row += 1;

                    cell = mae.DrawListRight(cell.Row, 1, new List<object>
                    {
                        "Masalah",
                        "Action",
                        "Stop",
                        "Start"
                    }, 1, 3, clsMagerExcel.BorderType.BorderAllThin);

                    cell.Row += 1;

                    cell = mae.DrawListRight(cell.Row, 1, new List<object>
                    {
                        "Masalah",
                        "Action",
                        "Stop",
                        "Start"
                    }, 2, 2, clsMagerExcel.BorderType.BorderAllThin);

                    cell.Row += 1;

                    cell = mae.DrawListRight(cell.Row, 1, new List<object>
                    {
                        "Masalah",
                        "Action",
                        "Stop",
                        "Start"
                    }, 1, 3);

                    cell.Row += 1;

                    cell = mae.DrawListRight(cell.Row, 1, new List<object>
                    {
                        "Masalah",
                        "Action",
                        "Stop",
                        "Start"
                    }, 2, 4, clsMagerExcel.BorderType.BorderAllThin, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    cell.Row += 1;

                    cell = mae.DrawListRight(cell.Row, 1, new List<object>
                    {
                        "Masalah",
                        "Action",
                        "Stop",
                        "Start"
                    }, 1, 6, clsMagerExcel.BorderType.BorderAllThin, ExcelVerticalAlignment.Center, ExcelHorizontalAlignment.Center);

                    Session["MagerExcel"] = excelPackage.GetAsByteArray();
                    if (Session["MagerExcel"] != null)
                    {
                        byte[] data = Session["MagerExcel"] as byte[];
                        return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MagerExcel_" + Convert.ToDateTime(DateTime.Now).ToString("yyyyMMdd") + ".xlsx");
                    }
                    else
                    {
                        return new EmptyResult();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return Json(e.Message, JsonRequestBehavior.AllowGet);
            }
        }

        //eof
    }
}