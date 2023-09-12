using UploadFile.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using Microsoft.AspNetCore.Authorization;
using ClosedXML.Excel;
using ExcelDataReader;
using System.IO;


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
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                FileStream stream = System.IO.File.Open(xlsxFilePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader;

                //1. Reading Excel file
                if (Path.GetExtension(xlsxFilePath).ToUpper() == ".XLS")
                {
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                var dataSet = excelReader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true // Use first row is ColumnName here :D
                    }
                });

                var nowCsvLines = System.IO.File.ReadAllLines(csvFilePath);
                for (int i = 1; i < nowCsvLines.Length; i++) // Skip header
                {
                    var nowCsvValues = nowCsvLines[i].Split(',');
                    if (nowCsvValues != null)
                    {
                        string gradeSymbol = nowCsvValues[5]; // Assuming Grade Symbol is at index 5
                        string nNumber = nowCsvValues[4]; // Assuming N Number is at index 4


                        DataRow matchingRow = dataSet.Tables[0].AsEnumerable().Where(r => ((string)r["Student ID"]).Contains(nNumber)).First();

                        if (matchingRow != null)
                        {
                            if (gradeSymbol.Contains("-"))
                            {
                                string[] gradeParts = gradeSymbol.Split('-');
                                string leftPart = gradeParts[0];

                                if (leftPart == "0")
                                {
                                    matchingRow["Grade"] = "ZERO"; // Assuming Grade is at index 10
                                }
                                else
                                {
                                    matchingRow["Grade"] = leftPart; // Assuming Grade is at index 10
                                }
                                string rightPart = gradeParts[1];

                                if (rightPart == "Capped" || rightPart == "DL")
                                {
                                    matchingRow["Grade Change Reason"] = "DL"; // Assuming Grade Change is at index 14
                                    matchingRow["Comment"] = ""; // Assuming Comment is at index 15
                                }
                                else if (rightPart == "NE" || rightPart == "NN" || rightPart == "NS" || rightPart == "NK")
                                {
                                    matchingRow["Grade Change Reason"] = ""; // Assuming Grade Change is at index 14
                                    matchingRow["Comment"] = rightPart.Substring(0, 2); // Assuming Comment is at index 15
                                }
                                else
                                {
                                    matchingRow["Grade Change Reason"] = ""; // Assuming Grade Change is at index 14
                                    matchingRow["Comment"] = rightPart; // Assuming Comment is at index 15
                                }
                            }
                            else
                            {
                                matchingRow["Grade"] = gradeSymbol; // Assuming Grade is at index 10
                            }
                        }
                    }
                }
                excelReader.Close();
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
