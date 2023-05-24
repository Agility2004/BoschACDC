using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace BoschACDC.Models
{
    public class BoschModel
    {
        [Key, Column(Order = 1)]
        public string DeclarationNum { get; set; }
        [Key, Column(Order = 2)]
        public int LineNum { get; set; }
        public string BrokerID { get; set; }
        public string BrokerName { get; set; }
        public string DeclarationType { get; set; }
        public string CustomerID { get; set; }
        public string CustomerName { get; set; }
        public string ImporterID { get; set; }
        public string ImporterName { get; set; }
        public string ImporterReferenceNum { get; set; }
        public string ConsigneeID { get; set; }
        public string ConsigneeName { get; set; }
        public string ImportCountry { get; set; }
        public string ArrivalDate { get; set; }
        public string ImportDate { get; set; }
        public string ReleaseDate { get; set; }
        public string DeliveryDate { get; set; }
        public string ModeOfTransport { get; set; }
        public string CarrierID { get; set; }
        public string CarrierName { get; set; }
        public string PortOfFiling { get; set; }
        public string CustomsOffice { get; set; }
        public decimal TotalDeclarationValue { get; set; }
        public string CurrencyCode { get; set; }
        public string TotalFees { get; set; }
        public string ProductNum { get; set; }
        public string ProductDesc { get; set; }
        public string StyleNum { get; set; }
        public string BusinessUnit { get; set; }
        public string BusinessDivision { get; set; }
        public string SupplierID { get; set; }
        public string SupplierName { get; set; }
        public string CountryOfOrigin { get; set; }
        public string ManufacturerID { get; set; }
        public string ManufacturerName { get; set; }
        public string InvoiceNum { get; set; }
        public string GrossWeight { get; set; }
        public decimal NetWeight { get; set; }
        public string WeightUOM { get; set; }
        public int TxnQty { get; set; }
        public string TxnQtyUOM { get; set; }
        public decimal UnitValue { get; set; }
        public decimal TotalLineValue { get; set; }
        public decimal TotalDutiableLineValue { get; set; }
        public string HsNum { get; set; }
        public string HsNum2 { get; set; }
        public string WCOHsNum { get; set; }
        public int RptQty { get; set; }
        public string RptQtyUOM { get; set; }
        public string AddlRptQty { get; set; }
        public string AddlRptQtyUOM { get; set; }
        public decimal AdValoremDutyRate { get; set; }
        public string SpecificRate { get; set; }
        public decimal LineDuty { get; set; }
        public string AddlLineDuty { get; set; }
        public string PreferenceCode1 { get; set; }
        public string PreferenceCode2 { get; set; }
        public decimal TotalLineVATAmt { get; set; }
        public decimal VATRate { get; set; }
        public string TotalLineExciseAmt { get; set; }
        public string TotalLineAddlIndirectTaxAmt { get; set; }
        public string ExportCountry { get; set; }
        public string ExportDate { get; set; }
        public string INCOTerms { get; set; }
        public string PortOfLading { get; set; }
        public string PortOfUnlading { get; set; }
        public string MasterBillOfLading { get; set; }
        public string HouseBillOfLading { get; set; }
        public string RelatedPartyFlag { get; set; }
        public string Fees { get; set; }

    }
}
