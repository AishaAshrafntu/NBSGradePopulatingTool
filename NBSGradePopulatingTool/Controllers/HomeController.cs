﻿using IronXL;
using UploadFile.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using Microsoft.AspNetCore.Authorization;
using ClosedXML.Excel;


namespace NBSGradePopulatingTool.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        /// <summary>
        /// Index
        /// </summary>,
        /// <returns></returns>
        public IActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Process Files
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        [HttpPost]
        public IActionResult ProcessFile(SingleFileModel file)
        {
            try
            {
                if (file.NowCsvFile == null || file.BannerXlsxFile == null)
                {
                    ViewBag.ErrorMessage = "Both files must be loaded.";
                    return View("Index");
                }

                // Create a temporary directory to store uploaded files

                string tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                Directory.CreateDirectory(tempDirectory);

                var filePath = Path.GetFullPath(file.NowCsvFile.FileName);
                string nowCsvFilePath = Path.Combine(tempDirectory, file.NowCsvFile.FileName);
                string bannerXlsxFilePath = Path.Combine(tempDirectory, file.BannerXlsxFile.FileName);

                using (var nowCsvStream = new FileStream(nowCsvFilePath, FileMode.Create))
                using (var bannerXlsxStream = new FileStream(bannerXlsxFilePath, FileMode.Create))
                {
                    file.NowCsvFile.CopyTo(nowCsvStream);
                    file.BannerXlsxFile.CopyTo(bannerXlsxStream);
                }
                string fileName = Path.GetFileNameWithoutExtension(bannerXlsxFilePath) + "-completed.xlsx";
                using (XLWorkbook wb = new XLWorkbook())
                {
                    DataTable dt = this.ProcessCsvAndXlsx(nowCsvFilePath, bannerXlsxFilePath).Tables[0];

                    wb.Worksheets.Add(dt);
                    using (MemoryStream stream = new MemoryStream())
                    {
                        wb.SaveAs(stream);
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["message"] = ex.Message;
                return RedirectToAction("GetReportError");
            }
        }

        /// <summary>
        /// Process Csv And Xlsx file
        /// </summary>
        /// <param name="csvFilePath"></param>
        /// <param name="xlsxFilePath"></param>
        public DataSet ProcessCsvAndXlsx(string csvFilePath, string xlsxFilePath)
        {
            try
            {
                WorkBook csvWorkbook = WorkBook.Load(csvFilePath);
                WorkBook xlsxWorkbook = WorkBook.Load(xlsxFilePath);

                var csvSheet = csvWorkbook.DefaultWorkSheet;
                var xlsxSheet = xlsxWorkbook.DefaultWorkSheet;

                var nowCsvLines = System.IO.File.ReadAllLines(csvFilePath);
                for (int i = 1; i < nowCsvLines.Length; i++) // Skip header
                {
                    var nowCsvValues = nowCsvLines[i].Split(',');
                    if (nowCsvValues != null)
                    {
                        string gradeSymbol = nowCsvValues[5]; // Assuming Grade Symbol is at index 5
                        string nNumber = nowCsvValues[4]; // Assuming N Number is at index 4

                        var matchingRow = xlsxSheet.Rows
                            .FirstOrDefault(x => x.Columns[3].Value.ToString() == nNumber);

                        if (matchingRow != null)
                        {
                            if (gradeSymbol.Contains("-"))
                            {
                                string[] gradeParts = gradeSymbol.Split('-');
                                string leftPart = gradeParts[0];



                                if (leftPart == "0")
                                {
                                    matchingRow.Columns[10].Value = "ZERO"; // Assuming Grade is at index 10
                                }
                                else
                                {
                                    matchingRow.Columns[10].Value = leftPart; // Assuming Grade is at index 10
                                }
                                string rightPart = gradeParts[1];



                                if (rightPart == "Capped" || rightPart == "DL")
                                {
                                    matchingRow.Columns[14].Value = "DL"; // Assuming Grade Change is at index 14
                                    matchingRow.Columns[15].Value = ""; // Assuming Comment is at index 15
                                }
                                else if (rightPart == "NE" || rightPart == "NN" || rightPart == "NS" || rightPart == "NK")
                                {
                                    matchingRow.Columns[14].Value = ""; // Assuming Grade Change is at index 14
                                    matchingRow.Columns[15].Value = rightPart.Substring(0, 2); // Assuming Comment is at index 15
                                }
                                else
                                {
                                    matchingRow.Columns[14].Value = ""; // Assuming Grade Change is at index 14
                                    matchingRow.Columns[15].Value = rightPart; // Assuming Comment is at index 15
                                }
                            }
                            else
                            {
                                matchingRow.Columns[10].Value = gradeSymbol; // Assuming Grade is at index 10
                            }
                        }
                    }
                }
                System.Data.DataSet dataSet = xlsxWorkbook.ToDataSet(true); // Allow easy integration with DataGrids, SQL and EF
                return dataSet;
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        /// <summary>
        /// GetReportError
        /// </summary>,
        /// <returns></returns>
        public IActionResult GetReportError()
        {
            return View();
        }

    }
}
