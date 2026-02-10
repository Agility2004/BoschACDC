using BoschACDC.Class;
using BoschACDC.Data;
using BoschACDC.Models;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace BoschACDC.Controllers
{
    public class HomeController : Controller
    {
        private readonly BoschDbExportContext dbBoschExport;
        private readonly BoschDbImportContext dbBoschImport;

        #region "All Functions"
        private (DataTable, DataTable) SplitDataTableByDeliveryDate(DataTable dtBOSCH)
        {
            DataRow[] dr0409 = dtBOSCH.Select("DeliveryDate not is null");
            DataRow[] drNon0409 = dtBOSCH.Select("DeliveryDate is null");
            DataTable dt0409 = dr0409.CopyToDataTable();
            DataTable dtNon0409 = drNon0409.CopyToDataTable();

            return (dt0409, dtNon0409);
        }

        private MemoryStream ConvertDataTableToMemoryStream(DataTable dataTable)
        {
            var csvContent = BuildCSVContent(dataTable);
            var byteArray = Encoding.ASCII.GetBytes(csvContent);
            return new MemoryStream(byteArray);
        }

        private void AddEntryToArchive(ZipArchive archive, MemoryStream memoryStream, string database, string cmid, string suffix)
        {
            string strSuffix = (suffix == string.Empty ? "" : suffix);
            string fileName = $"{database}-{cmid}-{DateTime.Now:ddMMyy_HHmmss}{strSuffix}.csv";
            var entry = archive.CreateEntry(fileName);
            using (var entryStream = entry.Open())
            {
                memoryStream.CopyTo(entryStream);
            }
        }

        private MemoryStream CreateZipArchive(MemoryStream memoryStream0409, MemoryStream memoryStreamNon0409, string database, string cmid)
        {
            var zipStream = new MemoryStream();
            using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create, true))
            {
                AddEntryToArchive(archive, memoryStream0409, database, cmid, "");
                AddEntryToArchive(archive, memoryStreamNon0409, database, cmid, "--Non0409");
            }
            memoryStream0409.Seek(0, SeekOrigin.Begin);
            memoryStreamNon0409.Seek(0, SeekOrigin.Begin);
            zipStream.Seek(0, SeekOrigin.Begin);

            return zipStream;
        }
        #endregion

        public DataTable ConvertCsvToDataTable(IFormFile csvFile)
        {
            DataTable dataTable = new DataTable();
            List<BoschModel> lstBoschs = new List<BoschModel>();

            using (var stream = csvFile.OpenReadStream())
            using (var reader = new StreamReader(stream))
            {
                bool headersRead = false;
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split('|');

                    if (!headersRead)
                    {
                        foreach (var header in values)
                        {
                            dataTable.Columns.Add(header);
                        }
                        headersRead = true;
                    }
                    else
                    {
                        if (values.Length != dataTable.Columns.Count)
                        {
                            continue;
                        }

                        var row = dataTable.NewRow();
                        for (int i = 0; i < values.Length; i++)
                        {
                            row[i] = values[i];
                        }
                        dataTable.Rows.Add(row);
                    }
                }
            }

            return dataTable;
        }

        private DataTable ConvertExcelToDataTable(IFormFile file)
        {
            DataTable dtResult = new DataTable();

            if (file != null)
            {
                using (var stream = file.OpenReadStream())
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                    ExcelPackage package = new ExcelPackage();
                    package.Load(stream);
                    if (package.Workbook.Worksheets.Count > 0)
                    {
                        var hasHeader = true;
                        byte intStartRow = 1;
                        using (ExcelWorksheet workSheet = package.Workbook.Worksheets.First())
                        {
                            foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
                            {
                                if (string.IsNullOrEmpty(firstRowCell.Text.Trim()))
                                {
                                    intStartRow = 2;
                                    continue;
                                }
                            }

                            foreach (var firstRowCell in workSheet.Cells[intStartRow, 1, 1, workSheet.Dimension.End.Column])
                            {
                                dtResult.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                            }

                            ++intStartRow;
                            for (int rowNum = intStartRow; rowNum <= workSheet.Dimension.End.Row; rowNum++)
                            {
                                var wsRow = workSheet.Cells[rowNum, 1, rowNum, workSheet.Dimension.End.Column];
                                DataRow row = dtResult.Rows.Add();
                                foreach (var cell in wsRow) row[cell.Start.Column - 1] = cell.Text;
                            }
                        }
                    }
                }
            }

            return dtResult;
        }

        private DataTable DistinctProductCode(DataTable dtResult)
        {
            DataTable dtDistinct = new DataTable();
            dtDistinct = dtResult.DefaultView.ToTable(true, "ProductNum", "BusinessUnit");
            return dtDistinct;
        }

        private List<CompareModel> CompareProductCode(DataTable dtExcel, DataTable dtCsv)
        {
            List<CompareModel> lstCompare = new List<CompareModel>();

            foreach (DataRow drExcel in dtExcel.Rows)
            {
                string strProductNum = drExcel["ProductNum"].ToString();
                string strBusinessUnit = drExcel["BusinessUnit"].ToString();
                DataRow[] drCompare = dtCsv.Select("ProductNum = '" + strProductNum + "' AND BusinessUnit = '" + drExcel["BusinessUnit"].ToString() + "'");
                if (drCompare.Length == 0)
                {
                    CompareModel compare = new CompareModel();
                    compare.ProductNum = strProductNum;
                    compare.BusinessUnit = strBusinessUnit;
                    lstCompare.Add(compare);
                }
            }

            return lstCompare;
        }

        public HomeController(BoschDbExportContext _dbBoschExport, BoschDbImportContext _dbBoschImport)
        {
            dbBoschExport = _dbBoschExport;
            dbBoschImport = _dbBoschImport;
        }

        public IActionResult UpdateBU()
        {
            return View();
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult NoFileProvided()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public ActionResult GetDataBusinessUnit(string database, string cmid, string startDate, string stopDate, string subCode)
        {
            string[] arrCMID = (cmid == "ALL" ? new string[] { "BOSCH", "RBTY", "ROBOSCH", "REXROTH" } : arrCMID = new string[] { cmid });
            DataTable dtBOSCH = GetBoschData(database, arrCMID, startDate, stopDate, subCode, "");

            if (dtBOSCH.Rows.Count == 0) return Json(new { success = true, message = "Not found data" });

            if (database == "I")
            {
                DataRow[] drRow = dtBOSCH.Select("BusinessUnit = ''");
                if (drRow.Length == 0)
                {
                    return Json(new { success = true, message = "Successfully" });
                }
                else
                {
                    DataTable dtBU = drRow.CopyToDataTable();
                    List<BoschModel> lstBosch = BuildBoschModels(dtBU);
                    return Json(new { success = true, message = "Successfully", lstBosch = lstBosch });
                }
            }
            else
            {
                return Json(new { success = true, message = "Successfully" });
            }
        }

        private List<BoschModel> BuildBoschModels(DataTable dataTable)
        {
            List<BoschModel> boschModels = new List<BoschModel>();
            foreach (DataRow row in dataTable.Rows)
            {
                BoschModel model = new BoschModel
                {
                    DeclarationNum = row["DeclarationNum"].ToString(),
                    LineNum = Convert.ToInt32(row["LineNum"]),
                    ProductNum = row["ProductNum"].ToString(),
                    BusinessUnit = string.IsNullOrEmpty(row["BusinessUnit"].ToString()) ? "N/A" : row["BusinessUnit"].ToString()
                };
                boschModels.Add(model);
            }
            return boschModels;
        }

        [HttpPost]
        public JsonResult UpdateBU(List<string> lstBU)
        {
            string sql = "EXEC USP_UPDATE_DATA_BOSCH_ACDC @PRODUCT_CODE, @BUSINESS_UNIT";

            foreach (var item in lstBU)
            {
                string[] arrBU = item.Split("|");
                List<SqlParameter> para = new List<SqlParameter>
                {
                    new SqlParameter{ParameterName="@PRODUCT_CODE", Value = arrBU[0].ToString().Trim()},
                    new SqlParameter{ ParameterName="@BUSINESS_UNIT", Value = arrBU[1].ToString().Trim()}
                };
                dbBoschImport.Database.ExecuteSqlRaw(sql, para);
            }

            return Json("Ok");
        }

        [HttpGet]
        public IActionResult ExportToCSV(string database, string cmid, string startDate, string stopDate, string subCode, string boschs, bool excludeStatus, string decNo)
        {
            try
            {
                var decodedBoschs = Uri.UnescapeDataString(boschs);
                var lstBoschs = JsonConvert.DeserializeObject<List<string>>(decodedBoschs);

                //string[] arrCMID = (cmid == "ALL" ? new string[] { "BOSCH", "RBTY", "ROBOSCH", "REXROTH" } : arrCMID = new string[] { cmid });
                string[] arrCMID = (cmid == "ALL" ? new string[] { "BR", "ROBOSCH", "REXROTH" } : arrCMID = new string[] { cmid });
                string exclude = (excludeStatus ? "Y" : "N");

                DataTable dtBOSCH = GetBoschData(database, arrCMID, startDate, stopDate, subCode, decNo);

                if (dtBOSCH.Rows.Count == 0) return RedirectToAction("NoFileProvided");

                foreach (var item in lstBoschs)
                {
                    string[] data = item.Split('|');
                    string productNum = data[0];
                    string businessUnit = data[1];

                    foreach (DataRow row in dtBOSCH.Rows)
                    {
                        if (row["ProductNum"].ToString() == productNum) row["BusinessUnit"] = businessUnit;
                    }
                }

                if (exclude == "Y")
                {
                    var (dt0409, dtNon0409) = SplitDataTableByDeliveryDate(dtBOSCH);
                    var memoryStream0409 = ConvertDataTableToMemoryStream(dt0409);
                    var memoryStreamNon0409 = ConvertDataTableToMemoryStream(dtNon0409);
                    var zipStream = CreateZipArchive(memoryStream0409, memoryStreamNon0409, database, cmid);

                    return File(zipStream.ToArray(), "application/zip", "files.zip");
                }
                else
                {
                    var csvContent = BuildCSVContent(dtBOSCH);
                    var byteArray = Encoding.ASCII.GetBytes(csvContent);
                    var memoryStream = new MemoryStream(byteArray);
                    string fileName = $"{database}-{cmid}-{DateTime.Now:ddMMyy_HHmmss}.csv";
                    return File(memoryStream.ToArray(), "text/csv", fileName);
                }
            }
            catch (Exception ex)
            {
                return RedirectToAction("Index");
            }
        }

        private DataTable GetBoschData(string database, string[] arrCMID, string startDate, string stopDate, string subCode, string decNo)
        {
            try
            {
                DataTable dtBosch = new DataTable();
                int intTimeOut = 180;

                if (database == "E")
                {
                    string sql = "EXEC USP_SELECT_DATA_BOSCH_ACDC @CMID,@START_DATE,@STOP_DATE,@SUB_CODE";
                    foreach (var cmid in arrCMID)
                    {
                        var cmidList = cmid == "BR"
                            ? new[] { "BOSCH", "RBTY" }
                            : new[] { cmid };

                        foreach (var realCMID in cmidList)
                        {
                            var parameters = new List<SqlParameter>
                            {
                                new SqlParameter("@CMID", realCMID),
                                new SqlParameter("@START_DATE", startDate.Replace("-", "")),
                                new SqlParameter("@STOP_DATE", stopDate.Replace("-", "")),
                                new SqlParameter("@SUB_CODE", string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode),
                            };

                            var data = dbBoschExport.Boschs.FromSqlRaw<BoschModel>(sql, parameters.ToArray()).ToList();
                            dtBosch.Merge(ListtoDataTableConverter.ToDataTable(data));
                        }
                        //var parameters = new List<SqlParameter>
                        //{
                        //    new SqlParameter("@CMID", cmid),
                        //    new SqlParameter("@START_DATE", startDate.Replace("-", "")),
                        //    new SqlParameter("@STOP_DATE", stopDate.Replace("-", "")),
                        //    new SqlParameter("@SUB_CODE", string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode),
                        //};

                        //var data = dbBoschExport.Boschs.FromSqlRaw<BoschModel>(sql, parameters.ToArray()).ToList();
                        //dtBosch.Merge(ListtoDataTableConverter.ToDataTable(data));
                    }
                }
                else
                {
                    string sql = "EXEC USP_SELECT_DATA_BOSCH_ACDC @CMID,@START_DATE,@STOP_DATE,@SUB_CODE";
                    foreach (var cmid in arrCMID)
                    {
                        var cmidList = cmid == "BR"
                            ? new[] { "BOSCH", "RBTY" }
                            : new[] { cmid };

                        foreach (var realCMID in cmidList)
                        {
                            var parameters = new List<SqlParameter>
                            {
                                new SqlParameter("@CMID", realCMID),
                                new SqlParameter("@START_DATE", startDate.Replace("-", "")),
                                new SqlParameter("@STOP_DATE", stopDate.Replace("-", "")),
                                new SqlParameter("@SUB_CODE", string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode)
                            };

                            var data = dbBoschImport.Boschs.FromSqlRaw<BoschModel>(sql, parameters.ToArray()).ToList();
                            dtBosch.Merge(ListtoDataTableConverter.ToDataTable(data));
                        }

                        //var parameters = new List<SqlParameter>
                        //{
                        //    new SqlParameter("@CMID", cmid),
                        //    new SqlParameter("@START_DATE", startDate.Replace("-", "")),
                        //    new SqlParameter("@STOP_DATE", stopDate.Replace("-", "")),
                        //    new SqlParameter("@SUB_CODE", string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode)
                        //};


                        //var data = dbBoschImport.Boschs.FromSqlRaw<BoschModel>(sql, parameters.ToArray()).ToList();
                        //dtBosch.Merge(ListtoDataTableConverter.ToDataTable(data));
                    }

                    if (!string.IsNullOrEmpty(decNo))
                    {
                        if (decNo.EndsWith(",")) decNo = decNo.TrimEnd(',');

                        sql = "EXEC USP_SELECT_DATA_BOSCH_ACDC_ACROSS @CMID,@SUB_CODE,@DEC_NO";

                        foreach (var cmid in arrCMID)
                        {
                            var cmidList = cmid == "BR"
                                ? new[] { "BOSCH", "RBTY" }
                                : new[] { cmid };

                            foreach (var realCMID in cmidList)
                            {
                                var parameters = new List<SqlParameter>
                                {
                                    new SqlParameter("@CMID", realCMID),
                                    new SqlParameter("@SUB_CODE", string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode),
                                    new SqlParameter("@DEC_NO", string.IsNullOrEmpty(decNo) ? DBNull.Value : decNo)
                                };

                                using (var transaction = dbBoschImport.Database.BeginTransaction())
                                {
                                    dbBoschImport.Database.SetCommandTimeout(intTimeOut);
                                    var data = dbBoschImport.Boschs.FromSqlRaw<BoschModel>(sql, parameters.ToArray()).ToList();
                                    dtBosch.Merge(ListtoDataTableConverter.ToDataTable(data));
                                }
                            }

                            //var parameters = new List<SqlParameter>
                            //{
                            //    new SqlParameter("@CMID", cmid),
                            //    new SqlParameter("@SUB_CODE", string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode),
                            //    new SqlParameter("@DEC_NO", string.IsNullOrEmpty(decNo) ? DBNull.Value : decNo)
                            //};

                            //using (var transaction = dbBoschImport.Database.BeginTransaction())
                            //{
                            //    dbBoschImport.Database.SetCommandTimeout(intTimeOut);
                            //    var data = dbBoschImport.Boschs.FromSqlRaw<BoschModel>(sql, parameters.ToArray()).ToList();
                            //    dtBosch.Merge(ListtoDataTableConverter.ToDataTable(data));
                            //}
                        }
                    }
                }

                return dtBosch;
            }
            catch (Exception ex)
            {

                throw new Exception(ex.ToString());
            }
        }

        private string BuildCSVContent(DataTable dataTable)
        {
            var strBuilder = new StringBuilder();

            // Header
            var headers = string.Join("|", dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName));
            strBuilder.AppendLine(headers);

            // Data rows
            foreach (DataRow row in dataTable.Rows)
            {
                var rowData = string.Join("|", row.ItemArray.Select(item => item?.ToString() ?? string.Empty));
                strBuilder.AppendLine(rowData);
            }

            return strBuilder.ToString();
        }
    
        public IActionResult UploadExcelFile()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Upload(IFormFile excelFile, IFormFile csvFile)
        {
            if (excelFile != null && excelFile.Length > 0 && csvFile != null && csvFile.Length > 0)
            {
                DataTable dtExcel = ConvertExcelToDataTable(excelFile);
                DataTable dtCsv = ConvertCsvToDataTable(csvFile);
                DataTable dtExcelDistinct = DistinctProductCode(dtExcel);
                DataTable dtCsvDistinct = DistinctProductCode(dtCsv);
                List<CompareModel> lstCompare = CompareProductCode(dtExcelDistinct, dtCsvDistinct);

                if (lstCompare.Count == 0) return Json(new { success = true, message = "Not found data" });

                return Json(new { success = true, message = "Found data has changed", lstCompare = lstCompare });
            }

            return Json(new { success = true, message = "Found data has changed" });
        }
    }
}
