﻿using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System;
using System.Data;
using System.IO;

namespace BoschACDC.Controllers
{
    public class BillingController : Controller
    {
        private readonly IConfiguration config;

        public BillingController(IConfiguration _config)
        {
            config = _config;
        }
        public IActionResult Index()
        {
            return View();
        }

        private DataSet GetDataReimbursement(string startDate, string stopDate)
        {
            DataSet dsResult = new DataSet();

            try
            {
                string strConnection = config.GetValue<string>("ConnectionStrings:SysFreight");
                using (SqlDataAdapter myDataAdapter = new SqlDataAdapter("_REX_BILL_LCC_RBTH", strConnection))
                {
                    myDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                    myDataAdapter.SelectCommand.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = startDate;
                    myDataAdapter.SelectCommand.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = stopDate;
                    myDataAdapter.Fill(dsResult);
                }

                return dsResult;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private MemoryStream CreateExcelFile(DataSet dsResult)
        {
            try
            {
                DataTable dtSheet1 = dsResult.Tables[0];
                DataTable dtSheet2 = dsResult.Tables[1];

                string strTemplateReimbursement = config.GetValue<string>("AppSettings:PathTemplateReimbursement");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                MemoryStream stream = new MemoryStream();

                using (var package = new ExcelPackage(strTemplateReimbursement))
                {
                    var sheet = package.Workbook.Worksheets["Import_template"];
                    Int16 intRow = 11;

                    foreach (DataRow drRow in dtSheet1.Rows)
                    {
                        sheet.Cells[$"A{intRow.ToString()}"].Value = drRow["Row#"];
                        sheet.Cells[$"B{intRow.ToString()}"].Value = drRow["Entity_LSP_Billing_period"].ToString().Trim();
                        sheet.Cells[$"C{intRow.ToString()}"].Value = drRow["InvoiceDate"].ToString().Trim();
                        sheet.Cells[$"D{intRow.ToString()}"].Value = drRow["AwbOrBlNo"].ToString().Trim();
                        sheet.Cells[$"E{intRow.ToString()}"].Value = drRow["Etd"].ToString().Trim();
                        sheet.Cells[$"F{intRow.ToString()}"].Value = drRow["ServiceType"].ToString().Trim();
                        sheet.Cells[$"G{intRow.ToString()}"].Value = drRow["Cost_type"].ToString().Trim();
                        sheet.Cells[$"H{intRow.ToString()}"].Value = drRow["Business_Case"].ToString().Trim();
                        sheet.Cells[$"I{intRow.ToString()}"].Value = drRow["Freight_Mode"].ToString().Trim();
                        sheet.Cells[$"J{intRow.ToString()}"].Value = drRow["Transport_Mode"].ToString().Trim();
                        sheet.Cells[$"L{intRow.ToString()}"].Value = drRow["ShipperName"].ToString().Trim();
                        sheet.Cells[$"M{intRow.ToString()}"].Value = drRow["OriginName"].ToString().Trim();
                        sheet.Cells[$"N{intRow.ToString()}"].Value = drRow["CountryOfOrigin"].ToString().Trim();
                        sheet.Cells[$"P{intRow.ToString()}"].Value = drRow["Port_of_departure"].ToString().Trim();
                        sheet.Cells[$"Q{intRow.ToString()}"].Value = drRow["ConsigneeName"].ToString().Trim();
                        sheet.Cells[$"R{intRow.ToString()}"].Value = drRow["Destination_City"].ToString().Trim();
                        sheet.Cells[$"S{intRow.ToString()}"].Value = drRow["Destination_Country"].ToString().Trim();
                        sheet.Cells[$"U{intRow.ToString()}"].Value = drRow["DestCode"].ToString().Trim();
                        sheet.Cells[$"V{intRow.ToString()}"].Value = drRow["DeliveryType"].ToString().Trim();
                        sheet.Cells[$"W{intRow.ToString()}"].Value = drRow["NumberofContainer"];
                        sheet.Cells[$"X{intRow.ToString()}"].Value = drRow["GrossWeight"];
                        sheet.Cells[$"Y{intRow.ToString()}"].Value = drRow["ChargeWeight"];
                        sheet.Cells[$"Z{intRow.ToString()}"].Value = drRow["Volume"];
                        sheet.Cells[$"AA{intRow.ToString()}"].Value = drRow["Eta"].ToString().Trim();
                        sheet.Cells[$"AC{intRow.ToString()}"].Value = drRow["LocalAmt"];
                        sheet.Cells[$"AD{intRow.ToString()}"].Value = drRow["VatAmt"];
                        sheet.Cells[$"AH{intRow.ToString()}"].Value = drRow["InvoiceLocalAmt"];
                        sheet.Cells[$"AI{intRow.ToString()}"].Value = drRow["TotalVatAmt"];
                        sheet.Cells[$"AJ{intRow.ToString()}"].Value = drRow["InvoiceAndVatLocalAmt"];
                        sheet.Cells[$"AP{intRow.ToString()}"].Value = drRow["JobNo"].ToString().Trim();
                        sheet.Cells[$"BD{intRow.ToString()}"].Value = drRow["InvoiceNo"].ToString().Trim();
                        sheet.Cells[$"BE{intRow.ToString()}"].Value = drRow["R_Control"].ToString().Trim();
                        ++intRow;
                    }

                    intRow = 3;
                    var sheet2 = package.Workbook.Worksheets["Reimbursement_MOL"];
                    foreach (DataRow drRow in dtSheet2.Rows)
                    {
                        sheet2.Cells[$"D{intRow.ToString()}"].Value = drRow["LocalAmt"];
                        sheet2.Cells[$"E{intRow.ToString()}"].Value = drRow["Description"].ToString().Trim();
                        sheet2.Cells[$"F{intRow.ToString()}"].Value = drRow["Business_Case"].ToString().Trim();
                        sheet2.Cells[$"G{intRow.ToString()}"].Value = drRow["Trasport_Mode"].ToString().Trim();
                        sheet2.Cells[$"H{intRow.ToString()}"].Value = drRow["AwbOrBlNo"].ToString().Trim();
                        ++intRow;
                    }

                    package.SaveAs(stream);
                    stream.Position = 0;
                }
                return stream;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }

        [HttpGet]
        public IActionResult ExportToExcel(string startDate, string stopDate)
        {
            startDate = startDate.Replace("-", "").Trim();
            stopDate = stopDate.Replace("-", "").Trim();
            DataSet dsResult = GetDataReimbursement(startDate, stopDate);
            if (dsResult.Tables.Count == 2)
            {
                MemoryStream stream = CreateExcelFile(dsResult);
                string excelName = $"Billing-{DateTime.Now.ToString("yyyyMMdd-HHmmss")}.xlsx";
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            }
            return View();
        }

    }
}
