using BoschACDC.Class;
using BoschACDC.Data;
using BoschACDC.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                StringBuilder strBuilder = new StringBuilder();
                var row = string.Empty;
                row = "DeclarationNum|LineNum|BrokerID|BrokerName|DeclarationType|CustomerID|CustomerName|ImporterID|ImporterName|ImporterReferenceNum|ConsigneeID|ConsigneeName|ImportCountry|ArrivalDate|ImportDate|ReleaseDate|DeliveryDate|ModeOfTransport|CarrierID|CarrierName|PortOfFiling|CustomsOffice|TotalDeclarationValue|CurrencyCode|TotalFees|ProductNum|ProductDesc|StyleNum|BusinessUnit|BusinessDivision|SupplierID|SupplierName|CountryOfOrigin|ManufacturerID|ManufacturerName|InvoiceNum|GrossWeight|NetWeight|WeightUOM|TxnQty|TxnQtyUOM|UnitValue|TotalLineValue|TotalDutiableLineValue|HsNum|HsNum2|WCOHsNum|RptQty|RptQtyUOM|AddlRptQty|AddlRptQtyUOM|AdValoremDutyRate|SpecificRate|LineDuty|AddlLineDuty|PreferenceCode1|PreferenceCode2|TotalLineVATAmt|VATRate|TotalLineExciseAmt|TotalLineAddlIndirectTaxAmt|ExportCountry|ExportDate|INCOTerms|PortOfLading|PortOfUnlading|MasterBillOfLading|HouseBillOfLading|RelatedPartyFlag|Fees";
                strBuilder.AppendLine(row);
                foreach (DataRow drRow in dtBOSCH.Rows)
                {
                    row = string.Empty;
                    foreach (DataColumn dcColumn in dtBOSCH.Columns) row += $"{ drRow[dcColumn].ToString() }{"|"}";
                    strBuilder.AppendLine(row.Substring(0, row.Length - 1));
                }

                var byteArray = Encoding.ASCII.GetBytes(strBuilder.ToString());
                var stream = new MemoryStream(byteArray);
                return File(stream.ToArray(), "text/csv", $"{database}-{ cmid }-{ DateTime.Now.ToString("ddMMyy_HHmmss") }.csv");
            }
            //return Content("No file name provided");
            return RedirectToAction("NoFileProvided");
        }
    }
}
