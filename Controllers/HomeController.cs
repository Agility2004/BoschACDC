using BoschACDC.Class;
using BoschACDC.Data;
using BoschACDC.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace BoschACDC.Controllers
{
    public class HomeController : Controller
    {
        private readonly BoschDbExportContext dbBoschExport;
        private readonly BoschDbImportContext dbBoschImport;


        public HomeController(BoschDbExportContext _dbBoschExport, BoschDbImportContext _dbBoschImport)
        {
            dbBoschExport = _dbBoschExport;
            dbBoschImport = _dbBoschImport;
        }

        public IActionResult UpdateBU()
        {
            return View();
        }

        //[AllowAnonymous]
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

        //[AllowAnonymous]
        public ActionResult getDataToCSV(string database, string cmid, string start_date, string stop_date, string subCode)
        {
            DataTable dtBOSCH = new DataTable();
            string[] arrCMID = new string[] { "BOSCH", "RBTY", "ROBOSCH" };
            string sql = "EXEC USP_SELECT_DATA_BOSCH_ACDC @CMID,@START_DATE,@STOP_DATE,@SUB_CODE";

            if (cmid == "ALL")
            {
                foreach (var CMID in arrCMID)
                {
                    DataTable dtResult = new DataTable();
                    List<SqlParameter> para = new List<SqlParameter>
                    {
                        new SqlParameter{ ParameterName = "@CMID", Value = CMID },
                        new SqlParameter{ ParameterName = "@START_DATE", Value = start_date.Replace("-","") },
                        new SqlParameter{ ParameterName = "@STOP_DATE", Value = stop_date.Replace("-","") },
                        new SqlParameter{ ParameterName = "@SUB_CODE", Value = (string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode) }
                    };

                    if (database == "E")
                    {
                        var data = dbBoschExport.Boschs.FromSqlRaw<BoschModel>(sql, para.ToArray()).ToList();
                        dtResult = ListtoDataTableConverter.ToDataTable(data);
                    }
                    else
                    {
                        var data = dbBoschImport.Boschs.FromSqlRaw<BoschModel>(sql, para.ToArray()).ToList();
                        dtResult = ListtoDataTableConverter.ToDataTable(data);
                    }

                    dtBOSCH.Merge(dtResult);
                }
            }
            else
            {
                List<SqlParameter> para = new List<SqlParameter>
                {
                    new SqlParameter{ ParameterName = "@CMID", Value = cmid },
                    new SqlParameter{ ParameterName = "@START_DATE", Value = start_date.Replace("-","") },
                    new SqlParameter{ ParameterName = "@STOP_DATE", Value = stop_date.Replace("-","") },
                    new SqlParameter{ ParameterName = "@SUB_CODE", Value = (string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode) }
                };

                if (database == "E")
                {
                    var data = dbBoschExport.Boschs.FromSqlRaw<BoschModel>(sql, para.ToArray()).ToList();
                    dtBOSCH = ListtoDataTableConverter.ToDataTable(data);
                }
                else
                {
                    var data = dbBoschImport.Boschs.FromSqlRaw<BoschModel>(sql, para.ToArray()).ToList();
                    dtBOSCH = ListtoDataTableConverter.ToDataTable(data);
                }
            }
            
            if (dtBOSCH.Rows.Count > 0)
            {
                if (database == "I")
                {
                    DataRow[] drRow = dtBOSCH.Select("BusinessUnit = ''");
                    if (drRow.Length == 0)
                    {
                        StringBuilder strBuilder = new StringBuilder();
                        var row = string.Empty;
                        row = "DeclarationNum|LineNum|BrokerID|BrokerName|DeclarationType|CustomerID|CustomerName|ImporterID|ImporterName|ImporterReferenceNum|ConsigneeID|ConsigneeName|ImportCountry|ArrivalDate|ImportDate|ReleaseDate|DeliveryDate|ModeOfTransport|CarrierID|CarrierName|PortOfFiling|CustomsOffice|TotalDeclarationValue|CurrencyCode|TotalFees|ProductNum|ProductDesc|StyleNum|BusinessUnit|BusinessDivision|SupplierID|SupplierName|CountryOfOrigin|ManufacturerID|ManufacturerName|InvoiceNum|GrossWeight|NetWeight|WeightUOM|TxnQty|TxnQtyUOM|UnitValue|TotalLineValue|TotalDutiableLineValue|HsNum|HsNum2|WCOHsNum|RptQty|RptQtyUOM|AddlRptQty|AddlRptQtyUOM|AdValoremDutyRate|SpecificRate|LineDuty|AddlLineDuty|PreferenceCode1|PreferenceCode2|TotalLineVATAmt|VATRate|TotalLineExciseAmt|TotalLineAddlIndirectTaxAmt|ExportCountry|ExportDate|INCOTerms|PortOfLading|PortOfUnlading|MasterBillOfLading|HouseBillOfLading|RelatedPartyFlag|Fees";
                        strBuilder.AppendLine(row);
                        foreach (DataRow drBosch in dtBOSCH.Rows)
                        {
                            row = string.Empty;
                            foreach (DataColumn dcColumn in dtBOSCH.Columns) row += $"{ drBosch[dcColumn].ToString() }{"|"}";
                            strBuilder.AppendLine(row.Substring(0, row.Length - 1));
                        }

                        var byteArray = Encoding.ASCII.GetBytes(strBuilder.ToString());
                        var stream = new MemoryStream(byteArray);
                        return File(stream.ToArray(), "text/csv", $"{database}-{ cmid }-{ DateTime.Now.ToString("ddMMyy_HHmmss") }.csv");
                    }
                    else
                    {
                        DataTable dtBU = drRow.CopyToDataTable();

                        List<BoschModel> lstBosch = new List<BoschModel>();
                        foreach (DataRow drBU in dtBU.Rows)
                        {
                            BoschModel model = new BoschModel();
                            model.DeclarationNum = drBU["DeclarationNum"].ToString();
                            model.LineNum = Convert.ToInt32(drBU["LineNum"]);
                            model.ProductNum = drBU["ProductNum"].ToString();
                            model.BusinessUnit = drBU["BusinessUnit"].ToString();
                            lstBosch.Add(model);
                        }

                        ViewBag.database = database;
                        ViewBag.cmid = cmid;
                        ViewBag.start_date = start_date;
                        ViewBag.stop_date = stop_date;
                        ViewBag.lock_control = "Y";
                        return View("Index", lstBosch);
                    }
                }
                else
                {
                    StringBuilder strBuilder = new StringBuilder();
                    var row = string.Empty;
                    row = "DeclarationNum|LineNum|BrokerID|BrokerName|DeclarationType|CustomerID|CustomerName|ImporterID|ImporterName|ImporterReferenceNum|ConsigneeID|ConsigneeName|ImportCountry|ArrivalDate|ImportDate|ReleaseDate|DeliveryDate|ModeOfTransport|CarrierID|CarrierName|PortOfFiling|CustomsOffice|TotalDeclarationValue|CurrencyCode|TotalFees|ProductNum|ProductDesc|StyleNum|BusinessUnit|BusinessDivision|SupplierID|SupplierName|CountryOfOrigin|ManufacturerID|ManufacturerName|InvoiceNum|GrossWeight|NetWeight|WeightUOM|TxnQty|TxnQtyUOM|UnitValue|TotalLineValue|TotalDutiableLineValue|HsNum|HsNum2|WCOHsNum|RptQty|RptQtyUOM|AddlRptQty|AddlRptQtyUOM|AdValoremDutyRate|SpecificRate|LineDuty|AddlLineDuty|PreferenceCode1|PreferenceCode2|TotalLineVATAmt|VATRate|TotalLineExciseAmt|TotalLineAddlIndirectTaxAmt|ExportCountry|ExportDate|INCOTerms|PortOfLading|PortOfUnlading|MasterBillOfLading|HouseBillOfLading|RelatedPartyFlag|Fees";
                    strBuilder.AppendLine(row);
                    foreach (DataRow drBosch in dtBOSCH.Rows)
                    {
                        row = string.Empty;
                        foreach (DataColumn dcColumn in dtBOSCH.Columns) row += $"{ drBosch[dcColumn].ToString() }{"|"}";
                        strBuilder.AppendLine(row.Substring(0, row.Length - 1));
                    }

                    var byteArray = Encoding.ASCII.GetBytes(strBuilder.ToString());
                    var stream = new MemoryStream(byteArray);
                    return File(stream.ToArray(), "text/csv", $"{database}-{ cmid }-{ DateTime.Now.ToString("ddMMyy_HHmmss") }.csv");
                }
            }
            return RedirectToAction("NoFileProvided");
        }

        //public List<BoschModel> UseJArrayParseInNewtonsoftJson(string boschs)
        //{
        //    using StreamReader reader = new(boschs);
        //    var json = reader.ReadToEnd();
        //    var jarray = JArray.Parse(json);
        //    List<BoschModel> teachers = new();
        //    foreach (var item in jarray)
        //    {
        //        BoschModel teacher = item.ToObject<BoschModel>();
        //        teachers.Add(teacher);
        //    }
        //    return teachers;
        //}

        [HttpPost]
        public JsonResult UpdateBU(List<string> lstBU)
        {
            string sql = "EXEC USP_UPDATE_DATA_BOSCH_ADCD @PRODUCT_CODE, @BUSINESS_UNIT";

            foreach (var item in lstBU)
            {
                string[] arrBU = item.Split(",");
                List<SqlParameter> para = new List<SqlParameter>
                {
                    new SqlParameter{ParameterName="@PRODUCT_CODE", Value = arrBU[0].ToString().Trim()},
                    new SqlParameter{ ParameterName="@BUSINESS_UNIT", Value = arrBU[1].ToString().Trim()}
                };
                dbBoschImport.Database.ExecuteSqlRaw(sql, para);
            }

            return Json("Ok");
        }

        //[AllowAnonymous]
        //[RequestFormLimits(MultipartBodyLengthLimit = 104857600)]
        public ActionResult ExportToCSV(string database, string cmid, string start_date, string stop_date, string subCode, string boschs)
        {
            //List<BoschModel> lstBoschs = new List<BoschModel>();
            //lstBoschs = JsonConvert.DeserializeObject<List<BoschModel>>(boschs);
            //List<string> lstBoschs = boschs.Split(',').ToList();
            List<string> lstBoschs = JsonConvert.DeserializeObject<List<string>>(boschs);
            DataTable dtBOSCH = new DataTable();
            string[] arrCMID = new string[] { "BOSCH", "RBTY", "ROBOSCH" };
            string sql = "EXEC USP_SELECT_DATA_BOSCH_ACDC @CMID,@START_DATE,@STOP_DATE,@SUB_CODE";

            if (cmid == "ALL")
            {
                foreach (var CMID in arrCMID)
                {
                    DataTable dtResult = new DataTable();
                    List<SqlParameter> para = new List<SqlParameter>
                    {
                        new SqlParameter{ ParameterName = "@CMID", Value = CMID },
                        new SqlParameter{ ParameterName = "@START_DATE", Value = start_date.Replace("-","") },
                        new SqlParameter{ ParameterName = "@STOP_DATE", Value = stop_date.Replace("-","") },
                        new SqlParameter{ ParameterName = "@SUB_CODE", Value = (string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode) }
                    };

                    if (database == "E")
                    {
                        var data = dbBoschExport.Boschs.FromSqlRaw<BoschModel>(sql, para.ToArray()).ToList();
                        dtResult = ListtoDataTableConverter.ToDataTable(data);
                    }
                    else
                    {
                        var data = dbBoschImport.Boschs.FromSqlRaw<BoschModel>(sql, para.ToArray()).ToList();
                        dtResult = ListtoDataTableConverter.ToDataTable(data);
                    }

                    dtBOSCH.Merge(dtResult);
                }
            }
            else
            {
                List<SqlParameter> para = new List<SqlParameter>
                {
                    new SqlParameter{ ParameterName = "@CMID", Value = cmid },
                    new SqlParameter{ ParameterName = "@START_DATE", Value = start_date.Replace("-","") },
                    new SqlParameter{ ParameterName = "@STOP_DATE", Value = stop_date.Replace("-","") },
                    new SqlParameter{ ParameterName = "@SUB_CODE", Value = (string.IsNullOrEmpty(subCode) ? DBNull.Value : subCode) }
                };

                if (database == "E")
                {
                    var data = dbBoschExport.Boschs.FromSqlRaw<BoschModel>(sql, para.ToArray()).ToList();
                    dtBOSCH = ListtoDataTableConverter.ToDataTable(data);
                }
                else
                {
                    var data = dbBoschImport.Boschs.FromSqlRaw<BoschModel>(sql, para.ToArray()).ToList();
                    dtBOSCH = ListtoDataTableConverter.ToDataTable(data);
                }
            }

            if (dtBOSCH.Rows.Count > 0)
            {
                foreach (var item in lstBoschs)
                {
                    string[] data = item.Split('|');

                    IEnumerable<DataRow> rows = dtBOSCH.Rows.Cast<DataRow>().Where(r => r["ProductNum"].ToString() == data[0]);
                    rows.ToList().ForEach(r => r.SetField("BusinessUnit", data[1]));
                }

                StringBuilder strBuilder = new StringBuilder();
                var row = string.Empty;
                row = "DeclarationNum|LineNum|BrokerID|BrokerName|DeclarationType|CustomerID|CustomerName|ImporterID|ImporterName|ImporterReferenceNum|ConsigneeID|ConsigneeName|ImportCountry|ArrivalDate|ImportDate|ReleaseDate|DeliveryDate|ModeOfTransport|CarrierID|CarrierName|PortOfFiling|CustomsOffice|TotalDeclarationValue|CurrencyCode|TotalFees|ProductNum|ProductDesc|StyleNum|BusinessUnit|BusinessDivision|SupplierID|SupplierName|CountryOfOrigin|ManufacturerID|ManufacturerName|InvoiceNum|GrossWeight|NetWeight|WeightUOM|TxnQty|TxnQtyUOM|UnitValue|TotalLineValue|TotalDutiableLineValue|HsNum|HsNum2|WCOHsNum|RptQty|RptQtyUOM|AddlRptQty|AddlRptQtyUOM|AdValoremDutyRate|SpecificRate|LineDuty|AddlLineDuty|PreferenceCode1|PreferenceCode2|TotalLineVATAmt|VATRate|TotalLineExciseAmt|TotalLineAddlIndirectTaxAmt|ExportCountry|ExportDate|INCOTerms|PortOfLading|PortOfUnlading|MasterBillOfLading|HouseBillOfLading|RelatedPartyFlag|Fees";
                strBuilder.AppendLine(row);
                foreach (DataRow drBosch in dtBOSCH.Rows)
                {
                    row = string.Empty;
                    foreach (DataColumn dcColumn in dtBOSCH.Columns) row += $"{ drBosch[dcColumn].ToString() }{"|"}";
                    strBuilder.AppendLine(row.Substring(0, row.Length - 1));
                }

                var byteArray = Encoding.ASCII.GetBytes(strBuilder.ToString());
                var stream = new MemoryStream(byteArray);
                return File(stream.ToArray(), "text/csv", $"{database}-{ cmid }-{ DateTime.Now.ToString("ddMMyy_HHmmss") }.csv");
            }

            return RedirectToAction("NoFileProvided");
        }
    }
}
