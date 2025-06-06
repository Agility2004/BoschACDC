using Microsoft.AspNetCore.Mvc;
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

        private DataSet GetDataReimbursement(string startDate, string stopDate, string storedName)
        {
            DataSet dsResult = new DataSet();

            try
            {
                string strConnection = config.GetValue<string>("ConnectionStrings:SysFreight");
                using (SqlDataAdapter myDataAdapter = new SqlDataAdapter(storedName, strConnection))
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
                throw new Exception(ex.ToString());
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
                        sheet.Cells[$"E{intRow.ToString()}"].Value = drRow["JobNo"].ToString().Trim();
                        sheet.Cells[$"F{intRow.ToString()}"].Value = drRow["Etd"].ToString().Trim();
                        sheet.Cells[$"G{intRow.ToString()}"].Value = drRow["Eta"].ToString().Trim();
                        sheet.Cells[$"H{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"I{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"J{intRow.ToString()}"].Value = drRow["ServiceType"].ToString().Trim();
                        sheet.Cells[$"K{intRow.ToString()}"].Value = drRow["Cost_type"].ToString().Trim();
                        sheet.Cells[$"L{intRow.ToString()}"].Value = drRow["Business_Case"].ToString().Trim();
                        sheet.Cells[$"M{intRow.ToString()}"].Value = drRow["Freight_Mode"].ToString().Trim();
                        sheet.Cells[$"N{intRow.ToString()}"].Value = drRow["Transport_Mode"].ToString().Trim();
                        sheet.Cells[$"O{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"P{intRow.ToString()}"].Value = drRow["ShipperName"].ToString().Trim();
                        sheet.Cells[$"Q{intRow.ToString()}"].Value = drRow["OriginName"].ToString().Trim();
                        sheet.Cells[$"R{intRow.ToString()}"].Value = drRow["CountryOfOrigin"].ToString().Trim();
                        sheet.Cells[$"S{intRow.ToString()}"].Value = drRow["Origin_postal_code"].ToString().Trim();
                        sheet.Cells[$"T{intRow.ToString()}"].Value = drRow["Port_of_departure"].ToString().Trim();
                        sheet.Cells[$"U{intRow.ToString()}"].Value = drRow["ConsigneeName"].ToString().Trim();
                        sheet.Cells[$"V{intRow.ToString()}"].Value = drRow["Destination_City"].ToString().Trim();
                        sheet.Cells[$"W{intRow.ToString()}"].Value = drRow["Destination_Country"].ToString().Trim();
                        sheet.Cells[$"X{intRow.ToString()}"].Value = drRow["Destination_postal_code"].ToString().Trim();
                        sheet.Cells[$"Y{intRow.ToString()}"].Value = drRow["DestCode"].ToString().Trim();
                        sheet.Cells[$"Z{intRow.ToString()}"].Value = drRow["DeliveryType"].ToString().Trim();
                        sheet.Cells[$"AA{intRow.ToString()}"].Value = drRow["NumberofContainer"];
                        sheet.Cells[$"AB{intRow.ToString()}"].Value = drRow["GrossWeight"];
                        sheet.Cells[$"AC{intRow.ToString()}"].Value = drRow["ChargeWeight"];
                        sheet.Cells[$"AD{intRow.ToString()}"].Value = drRow["Volume"];
                        sheet.Cells[$"AE{intRow.ToString()}"].Value = drRow["LocalAmt"];
                        sheet.Cells[$"AF{intRow.ToString()}"].Value = drRow["VatAmt"];
                        sheet.Cells[$"AG{intRow.ToString()}"].Value = drRow["Remark_position"];
                        sheet.Cells[$"AH{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AI{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AJ{intRow.ToString()}"].Value = drRow["InvoiceLocalAmt"];
                        sheet.Cells[$"AK{intRow.ToString()}"].Value = drRow["TotalVatAmt"];
                        sheet.Cells[$"AL{intRow.ToString()}"].Value = drRow["InvoiceAndVatLocalAmt"];
                        sheet.Cells[$"AM{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AN{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AO{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AP{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AQ{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AR{intRow.ToString()}"].Value = drRow["CommercialInvoiceNo"].ToString().Trim();
                        sheet.Cells[$"AS{intRow.ToString()}"].Value = drRow["R_contorl"].ToString().Trim();
                        sheet.Cells[$"AT{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AU{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AV{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AW{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AX{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AY{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"AZ{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"BA{intRow.ToString()}"].Value = drRow["Container_type"].ToString().Trim();
                        sheet.Cells[$"BB{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"BC{intRow.ToString()}"].Value = string.Empty;
                        sheet.Cells[$"BD{intRow.ToString()}"].Value = drRow["DecNo"].ToString().Trim();
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
                        sheet2.Cells[$"I{intRow.ToString()}"].Value = drRow["Remark_position"].ToString().Trim();
                        ++intRow;
                    }

                    package.SaveAs(stream);
                    stream.Position = 0;
                }
                //using (var package = new ExcelPackage(strTemplateReimbursement))
                //{
                //    var sheet = package.Workbook.Worksheets["Import_template"];
                //    Int16 intRow = 11;

                //    foreach (DataRow drRow in dtSheet1.Rows)
                //    {
                //        sheet.Cells[$"A{intRow.ToString()}"].Value = drRow["Row#"];
                //        sheet.Cells[$"B{intRow.ToString()}"].Value = drRow["Entity_LSP_Billing_period"].ToString().Trim();
                //        sheet.Cells[$"C{intRow.ToString()}"].Value = drRow["InvoiceDate"].ToString().Trim();
                //        sheet.Cells[$"D{intRow.ToString()}"].Value = drRow["AwbOrBlNo"].ToString().Trim();
                //        sheet.Cells[$"E{intRow.ToString()}"].Value = drRow["Etd"].ToString().Trim();
                //        sheet.Cells[$"F{intRow.ToString()}"].Value = drRow["ServiceType"].ToString().Trim();
                //        sheet.Cells[$"G{intRow.ToString()}"].Value = drRow["Cost_type"].ToString().Trim();
                //        sheet.Cells[$"H{intRow.ToString()}"].Value = drRow["Business_Case"].ToString().Trim();
                //        sheet.Cells[$"I{intRow.ToString()}"].Value = drRow["Freight_Mode"].ToString().Trim();
                //        sheet.Cells[$"J{intRow.ToString()}"].Value = drRow["Transport_Mode"].ToString().Trim();
                //        sheet.Cells[$"L{intRow.ToString()}"].Value = drRow["ShipperName"].ToString().Trim();
                //        sheet.Cells[$"M{intRow.ToString()}"].Value = drRow["OriginName"].ToString().Trim();
                //        sheet.Cells[$"N{intRow.ToString()}"].Value = drRow["CountryOfOrigin"].ToString().Trim();
                //        sheet.Cells[$"P{intRow.ToString()}"].Value = drRow["Port_of_departure"].ToString().Trim();
                //        sheet.Cells[$"Q{intRow.ToString()}"].Value = drRow["ConsigneeName"].ToString().Trim();
                //        sheet.Cells[$"R{intRow.ToString()}"].Value = drRow["Destination_City"].ToString().Trim();
                //        sheet.Cells[$"S{intRow.ToString()}"].Value = drRow["Destination_Country"].ToString().Trim();
                //        sheet.Cells[$"U{intRow.ToString()}"].Value = drRow["DestCode"].ToString().Trim();
                //        sheet.Cells[$"V{intRow.ToString()}"].Value = drRow["DeliveryType"].ToString().Trim();
                //        sheet.Cells[$"W{intRow.ToString()}"].Value = drRow["NumberofContainer"];
                //        sheet.Cells[$"X{intRow.ToString()}"].Value = drRow["GrossWeight"];
                //        sheet.Cells[$"Y{intRow.ToString()}"].Value = drRow["ChargeWeight"];
                //        sheet.Cells[$"Z{intRow.ToString()}"].Value = drRow["Volume"];
                //        sheet.Cells[$"AA{intRow.ToString()}"].Value = drRow["Eta"].ToString().Trim();
                //        sheet.Cells[$"AC{intRow.ToString()}"].Value = drRow["LocalAmt"];
                //        sheet.Cells[$"AD{intRow.ToString()}"].Value = drRow["VatAmt"];
                //        sheet.Cells[$"AH{intRow.ToString()}"].Value = drRow["InvoiceLocalAmt"];
                //        sheet.Cells[$"AI{intRow.ToString()}"].Value = drRow["TotalVatAmt"];
                //        sheet.Cells[$"AJ{intRow.ToString()}"].Value = drRow["InvoiceAndVatLocalAmt"];
                //        sheet.Cells[$"AP{intRow.ToString()}"].Value = drRow["JobNo"].ToString().Trim();
                //        sheet.Cells[$"BD{intRow.ToString()}"].Value = drRow["InvoiceNo"].ToString().Trim();
                //        sheet.Cells[$"BE{intRow.ToString()}"].Value = drRow["R_Control"].ToString().Trim();
                //        ++intRow;
                //    }

                //    intRow = 3;
                //    var sheet2 = package.Workbook.Worksheets["Reimbursement_MOL"];
                //    foreach (DataRow drRow in dtSheet2.Rows)
                //    {
                //        sheet2.Cells[$"D{intRow.ToString()}"].Value = drRow["LocalAmt"];
                //        sheet2.Cells[$"E{intRow.ToString()}"].Value = drRow["Description"].ToString().Trim();
                //        sheet2.Cells[$"F{intRow.ToString()}"].Value = drRow["Business_Case"].ToString().Trim();
                //        sheet2.Cells[$"G{intRow.ToString()}"].Value = drRow["Trasport_Mode"].ToString().Trim();
                //        sheet2.Cells[$"H{intRow.ToString()}"].Value = drRow["AwbOrBlNo"].ToString().Trim();
                //        sheet2.Cells[$"I{intRow.ToString()}"].Value = drRow["Remark"].ToString().Trim();
                //        ++intRow;
                //    }

                //    package.SaveAs(stream);
                //    stream.Position = 0;
                //}
                return stream;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }

        private MemoryStream CreateExcelFileNewTemplate(DataSet dsResult)
        {
            try
            {
                DataTable dtSheet1 = dsResult.Tables[0];
                //DataTable dtSheet2 = dsResult.Tables[1];

                string strTemplateReimbursement = config.GetValue<string>("AppSettings:PathNewTemplateReimbursement");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                MemoryStream stream = new MemoryStream();

                using (var package = new ExcelPackage(strTemplateReimbursement))
                {
                    var sheet = package.Workbook.Worksheets["B&O Lsp Billing"];
                    Int16 intRow = 3;

                    foreach (DataRow drRow in dtSheet1.Rows)
                    {
                        sheet.Cells[$"A{intRow.ToString()}"].Value = drRow["Row#"];
                        sheet.Cells[$"B{intRow.ToString()}"].Value = drRow["BillingPreiod"].ToString().Trim();
                        sheet.Cells[$"C{intRow.ToString()}"].Value = drRow["CustomerCode"].ToString().Trim();
                        sheet.Cells[$"D{intRow.ToString()}"].Value = drRow["LSP"].ToString().Trim();
                        sheet.Cells[$"E{intRow.ToString()}"].Value = drRow["InvoiceDate"].ToString().Trim();
                        sheet.Cells[$"F{intRow.ToString()}"].Value = drRow["AwbOrBlNo"].ToString().Trim();
                        sheet.Cells[$"G{intRow.ToString()}"].Value = drRow["DecNo"].ToString().Trim();
                        sheet.Cells[$"H{intRow.ToString()}"].Value = drRow["Declaration number 2 (SMK)"].ToString().Trim();
                        sheet.Cells[$"I{intRow.ToString()}"].Value = drRow["Declaration nunber 3 (AEO)"].ToString().Trim();
                        sheet.Cells[$"J{intRow.ToString()}"].Value = drRow["CommercialInvoiceNo"].ToString().Trim();
                        sheet.Cells[$"K{intRow.ToString()}"].Value = drRow["Reference number"].ToString().Trim();
                        sheet.Cells[$"L{intRow.ToString()}"].Value = drRow["IsDocumentation"].ToString().Trim();
                        sheet.Cells[$"M{intRow.ToString()}"].Value = drRow["Document Amendment"].ToString().Trim();
                        sheet.Cells[$"N{intRow.ToString()}"].Value = drRow["Document Amendment Time"].ToString().Trim();
                        sheet.Cells[$"O{intRow.ToString()}"].Value = drRow["Document Amendment Time"].ToString().Trim();
                        sheet.Cells[$"P{intRow.ToString()}"].Value = drRow["Consignment"].ToString().Trim();
                        sheet.Cells[$"Q{intRow.ToString()}"].Value = drRow["IsHandling"].ToString().Trim();
                        sheet.Cells[$"R{intRow.ToString()}"].Value = drRow["IsInspection"].ToString().Trim();
                        sheet.Cells[$"S{intRow.ToString()}"].Value = drRow["IsTransportation"].ToString().Trim();
                        sheet.Cells[$"T{intRow.ToString()}"].Value = drRow["COO application"].ToString().Trim();
                        sheet.Cells[$"U{intRow.ToString()}"].Value = drRow["DO release"].ToString().Trim();
                        sheet.Cells[$"V{intRow.ToString()}"].Value = drRow["Duty"].ToString().Trim();
                        sheet.Cells[$"W{intRow.ToString()}"].Value = drRow["Monthly Storage"].ToString().Trim();
                        sheet.Cells[$"X{intRow.ToString()}"].Value = drRow["Break bulk"].ToString().Trim();
                        sheet.Cells[$"Y{intRow.ToString()}"].Value = drRow["Bonded warehouse"].ToString().Trim();
                        sheet.Cells[$"Z{intRow.ToString()}"].Value = drRow["Hand over"].ToString().Trim();
                        sheet.Cells[$"Z{intRow.ToString()}"].Value = drRow["System Update"].ToString().Trim();
                        sheet.Cells[$"AA{intRow.ToString()}"].Value = drRow["Document input for 3rd Party"];
                        sheet.Cells[$"AB{intRow.ToString()}"].Value = drRow["IsLoading"];
                        sheet.Cells[$"AC{intRow.ToString()}"].Value = drRow["IsLoading"];
                        sheet.Cells[$"AD{intRow.ToString()}"].Value = drRow["Port drayage"];
                        sheet.Cells[$"AE{intRow.ToString()}"].Value = drRow["Forklift handling"];
                        sheet.Cells[$"AF{intRow.ToString()}"].Value = drRow["Waiting time"];
                        sheet.Cells[$"AG{intRow.ToString()}"].Value = drRow["Back invoice"];
                        sheet.Cells[$"AH{intRow.ToString()}"].Value = drRow["IsOther"];
                        sheet.Cells[$"AI{intRow.ToString()}"].Value = drRow["Surcharge 1"];
                        sheet.Cells[$"AJ{intRow.ToString()}"].Value = drRow["Surcharge 2"];
                        sheet.Cells[$"AK{intRow.ToString()}"].Value = drRow["Surcharge 3"];
                        sheet.Cells[$"AL{intRow.ToString()}"].Value = drRow["Surcharge 4"];
                        sheet.Cells[$"AM{intRow.ToString()}"].Value = drRow["Surcharge 5"];
                        sheet.Cells[$"AN{intRow.ToString()}"].Value = drRow["Surcharge 6"];
                        sheet.Cells[$"AO{intRow.ToString()}"].Value = drRow["Surcharge 7"];
                        sheet.Cells[$"AP{intRow.ToString()}"].Value = drRow["Surcharge 8"];
                        sheet.Cells[$"AQ{intRow.ToString()}"].Value = drRow["Surcharge 9"];
                        sheet.Cells[$"AR{intRow.ToString()}"].Value = drRow["Surcharge 10"];
                        sheet.Cells[$"AS{intRow.ToString()}"].Value = drRow["Business_Case"];
                        sheet.Cells[$"AT{intRow.ToString()}"].Value = drRow["Freight_Mode"];
                        sheet.Cells[$"AU{intRow.ToString()}"].Value = drRow["Transport_Mode"];
                        sheet.Cells[$"AV{intRow.ToString()}"].Value = drRow["Dangerous good"];
                        sheet.Cells[$"AW{intRow.ToString()}"].Value = drRow["Service type"];
                        sheet.Cells[$"AX{intRow.ToString()}"].Value = drRow["Container size"];
                        sheet.Cells[$"AY{intRow.ToString()}"].Value = drRow["Truck size"];
                        sheet.Cells[$"AZ{intRow.ToString()}"].Value = drRow["NumberofContainer"];
                        sheet.Cells[$"BA{intRow.ToString()}"].Value = drRow["Number of package"].ToString().Trim();
                        sheet.Cells[$"BB{intRow.ToString()}"].Value = drRow["Internal shipment ID"];
                        sheet.Cells[$"BC{intRow.ToString()}"].Value = drRow["Single item"];
                        sheet.Cells[$"BD{intRow.ToString()}"].Value = drRow["Sub item"].ToString().Trim();
                        sheet.Cells[$"BE{intRow.ToString()}"].Value = drRow["Total item"].ToString().Trim();
                        sheet.Cells[$"BF{intRow.ToString()}"].Value = drRow["ShipperName"].ToString().Trim();
                        sheet.Cells[$"BG{intRow.ToString()}"].Value = drRow["OriginName"].ToString().Trim();
                        sheet.Cells[$"BH{intRow.ToString()}"].Value = drRow["CountryOfOrigin"].ToString().Trim();
                        sheet.Cells[$"BI{intRow.ToString()}"].Value = drRow["Port_of_departure"].ToString().Trim();
                        sheet.Cells[$"BJ{intRow.ToString()}"].Value = drRow["ConsigneeName"].ToString().Trim();
                        sheet.Cells[$"BK{intRow.ToString()}"].Value = drRow["Destination_City"].ToString().Trim();
                        sheet.Cells[$"BL{intRow.ToString()}"].Value = drRow["Destination_Country"].ToString().Trim();
                        sheet.Cells[$"BM{intRow.ToString()}"].Value = drRow["DestCode"].ToString().Trim();
                        sheet.Cells[$"BN{intRow.ToString()}"].Value = drRow["DeliveryType"].ToString().Trim();
                        sheet.Cells[$"BO{intRow.ToString()}"].Value = drRow["GrossWeight"].ToString().Trim();
                        sheet.Cells[$"BP{intRow.ToString()}"].Value = drRow["ChargeWeight"].ToString().Trim();
                        sheet.Cells[$"BQ{intRow.ToString()}"].Value = drRow["Volume"].ToString().Trim();
                        sheet.Cells[$"BR{intRow.ToString()}"].Value = drRow["Pick up date"].ToString().Trim();
                        sheet.Cells[$"BS{intRow.ToString()}"].Value = drRow["Etd"].ToString().Trim();
                        sheet.Cells[$"BT{intRow.ToString()}"].Value = drRow["Eta"].ToString().Trim();
                        sheet.Cells[$"BU{intRow.ToString()}"].Value = drRow["DecDate"].ToString().Trim();
                        sheet.Cells[$"BV{intRow.ToString()}"].Value = drRow["CloseJobDate"].ToString().Trim();
                        sheet.Cells[$"BW{intRow.ToString()}"].Value = drRow["Storage day(s)"].ToString().Trim();
                        sheet.Cells[$"BX{intRow.ToString()}"].Value = drRow["Night time"].ToString().Trim();
                        sheet.Cells[$"BY{intRow.ToString()}"].Value = drRow["Weekend"].ToString().Trim();
                        sheet.Cells[$"BZ{intRow.ToString()}"].Value = drRow["HS code"].ToString().Trim();
                        sheet.Cells[$"CA{intRow.ToString()}"].Value = drRow["Sub Port"].ToString().Trim();
                        sheet.Cells[$"CB{intRow.ToString()}"].Value = drRow["Declaration Value"];
                        sheet.Cells[$"CC{intRow.ToString()}"].Value = drRow["Total CW"];
                        sheet.Cells[$"CD{intRow.ToString()}"].Value = drRow["PDCL"].ToString().Trim();
                        sheet.Cells[$"CE{intRow.ToString()}"].Value = drRow["Total Non Tax amount"].ToString().Trim();
                        sheet.Cells[$"CF{intRow.ToString()}"].Value = drRow["Previous shipping mode"].ToString().Trim();
                        sheet.Cells[$"CG{intRow.ToString()}"].Value = drRow["GB"].ToString().Trim();
                        sheet.Cells[$"CH{intRow.ToString()}"].Value = drRow["BU"].ToString().Trim();
                        sheet.Cells[$"CI{intRow.ToString()}"].Value = drRow["Istar"].ToString().Trim();
                        sheet.Cells[$"CJ{intRow.ToString()}"].Value = drRow["Vendor code"].ToString().Trim();
                        sheet.Cells[$"CK{intRow.ToString()}"].Value = drRow["Company code"].ToString().Trim();
                        sheet.Cells[$"CL{intRow.ToString()}"].Value = drRow["Goods type"].ToString().Trim();
                        sheet.Cells[$"CM{intRow.ToString()}"].Value = drRow["Declaration amount"].ToString().Trim();
                        sheet.Cells[$"CN{intRow.ToString()}"].Value = drRow["Amendment"].ToString().Trim();
                        sheet.Cells[$"CO{intRow.ToString()}"].Value = drRow["Handling amount"].ToString().Trim();
                        sheet.Cells[$"CP{intRow.ToString()}"].Value = drRow["Arrival Notice amount"].ToString().Trim();
                        sheet.Cells[$"CQ{intRow.ToString()}"].Value = drRow["Inspection fee"].ToString().Trim();
                        sheet.Cells[$"CR{intRow.ToString()}"].Value = drRow["Transportation amount"].ToString().Trim();
                        sheet.Cells[$"CS{intRow.ToString()}"].Value = drRow["COO application"].ToString().Trim();
                        sheet.Cells[$"CT{intRow.ToString()}"].Value = "";
                        sheet.Cells[$"CU{intRow.ToString()}"].Value = "";
                        sheet.Cells[$"CV{intRow.ToString()}"].Value = drRow["Storage amount"].ToString().Trim();
                        sheet.Cells[$"CW{intRow.ToString()}"].Value = "";
                        sheet.Cells[$"CX{intRow.ToString()}"].Value = "";
                        sheet.Cells[$"CY{intRow.ToString()}"].Value = "";
                        sheet.Cells[$"CZ{intRow.ToString()}"].Value = "";
                        sheet.Cells[$"DA{intRow.ToString()}"].Value = "";
                        sheet.Cells[$"DB{intRow.ToString()}"].Value = drRow["Document input for 3rd Party amount"];
                        sheet.Cells[$"DC{intRow.ToString()}"].Value = drRow["Post submission amount"];
                        sheet.Cells[$"DD{intRow.ToString()}"].Value = drRow["Loading/Unloading amount"].ToString().Trim();
                        sheet.Cells[$"DE{intRow.ToString()}"].Value = drRow["Port drayage amount"].ToString().Trim();
                        sheet.Cells[$"DF{intRow.ToString()}"].Value = drRow["Forklift handling amount"].ToString().Trim();
                        sheet.Cells[$"DG{intRow.ToString()}"].Value = drRow["Waiting time amount"].ToString().Trim();
                        sheet.Cells[$"DH{intRow.ToString()}"].Value = drRow["ReimbursementAmt"].ToString().Trim();
                        sheet.Cells[$"DI{intRow.ToString()}"].Value = drRow["Other amount"].ToString().Trim();
                        sheet.Cells[$"DJ{intRow.ToString()}"].Value = drRow["Surcharge 1"].ToString().Trim();
                        sheet.Cells[$"DK{intRow.ToString()}"].Value = drRow["Surcharge 2"].ToString().Trim();
                        sheet.Cells[$"DL{intRow.ToString()}"].Value = drRow["Surcharge 3"].ToString().Trim();
                        sheet.Cells[$"DM{intRow.ToString()}"].Value = drRow["Surcharge 4"].ToString().Trim();
                        sheet.Cells[$"DN{intRow.ToString()}"].Value = drRow["Surcharge 5"].ToString().Trim();
                        sheet.Cells[$"DO{intRow.ToString()}"].Value = drRow["Surcharge 6"].ToString().Trim();
                        sheet.Cells[$"DP{intRow.ToString()}"].Value = drRow["Surcharge 7"].ToString().Trim();
                        sheet.Cells[$"DQ{intRow.ToString()}"].Value = drRow["Surcharge 8"].ToString().Trim();
                        sheet.Cells[$"DR{intRow.ToString()}"].Value = drRow["Surcharge 9"].ToString().Trim();
                        sheet.Cells[$"DS{intRow.ToString()}"].Value = drRow["Surcharge 10"].ToString().Trim();
                        sheet.Cells[$"DT{intRow.ToString()}"].Value = drRow["Remark"].ToString().Trim();
                        sheet.Cells[$"DU{intRow.ToString()}"].Value = drRow["Remark 1"].ToString().Trim();
                        sheet.Cells[$"DV{intRow.ToString()}"].Value = drRow["Remark 2"].ToString().Trim();
                        sheet.Cells[$"DW{intRow.ToString()}"].Value = drRow["Remark 3"].ToString().Trim();
                        sheet.Cells[$"DX{intRow.ToString()}"].Value = drRow["Remark 4"].ToString().Trim();
                        sheet.Cells[$"DY{intRow.ToString()}"].Value = drRow["Remark 5"].ToString().Trim();
                        sheet.Cells[$"DZ{intRow.ToString()}"].Value = drRow["Remark 6"].ToString().Trim();
                        sheet.Cells[$"EA{intRow.ToString()}"].Value = drRow["Remark 7"].ToString().Trim();
                        sheet.Cells[$"EB{intRow.ToString()}"].Value = drRow["Remark 8"].ToString().Trim();
                        sheet.Cells[$"EC{intRow.ToString()}"].Value = drRow["Remark 9"].ToString().Trim();
                        sheet.Cells[$"ED{intRow.ToString()}"].Value = drRow["Remark 10"].ToString().Trim();
                        sheet.Cells[$"EE{intRow.ToString()}"].Value = drRow["Net amount"].ToString().Trim();
                        sheet.Cells[$"EF{intRow.ToString()}"].Value = drRow["Tax rate"].ToString().Trim();
                        sheet.Cells[$"EG{intRow.ToString()}"].Value = drRow["Tax amount"].ToString().Trim();
                        sheet.Cells[$"EH{intRow.ToString()}"].Value = drRow["Total amount"].ToString().Trim();
                        sheet.Cells[$"EI{intRow.ToString()}"].Value = drRow["CurrCode"].ToString().Trim();
                        sheet.Cells[$"EJ{intRow.ToString()}"].Value = drRow["PO number"].ToString().Trim();
                        sheet.Cells[$"EK{intRow.ToString()}"].Value = drRow["Credit note number"].ToString().Trim();
                        sheet.Cells[$"EL{intRow.ToString()}"].Value = drRow["Credit note date"].ToString().Trim();
                        sheet.Cells[$"EM{intRow.ToString()}"].Value = drRow["InvoiceNo"].ToString().Trim();
                        sheet.Cells[$"EN{intRow.ToString()}"].Value = drRow["Tax invoice date"].ToString().Trim();
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
        public IActionResult ExportToExcel(string startDate, string stopDate, string templateOption)
        {
            startDate = startDate.Replace("-", "").Trim();
            stopDate = stopDate.Replace("-", "").Trim();
            string excelName = $"Billing-{DateTime.Now.ToString("yyyyMMdd-HHmmss")}.xlsx";
            MemoryStream stream = null;

            if (templateOption == "oldTemplate")
            {
                DataSet dsResult = GetDataReimbursement(startDate, stopDate, "_REX_BILL_LCC_RBTH");
                if (dsResult.Tables.Count == 2)
                {
                    stream = CreateExcelFile(dsResult);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
                }
            }
            else
            {
                DataSet dsResult = GetDataReimbursement(startDate, stopDate, "_REX_BILL_LCC_RBTH_NEW");
                //if (dsResult.Tables.Count == 2)
                if (dsResult.Tables.Count == 1)
                {
                    stream = CreateExcelFileNewTemplate(dsResult);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
                }
            }

            if (stream != null)
            {
                Response.Headers.Add("Content-Disposition", $"attachment; filename={excelName}");
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            }

            return View();
        }

    }
}
