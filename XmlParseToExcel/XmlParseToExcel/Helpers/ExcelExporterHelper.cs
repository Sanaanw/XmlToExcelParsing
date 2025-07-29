using OfficeOpenXml;
using System.ComponentModel;
using System.IO;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace XmlParseToExcel.Helpers
{
    public class ExcelExporterHelper
    {
        public static void ExportXmlToExcel(OrderDto order, string outputPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Order");

                // Header
                worksheet.Cells[1, 1].Value = "Order Number";
                worksheet.Cells[1, 2].Value = order.Header.OrderNumber;
                worksheet.Cells[2, 1].Value = "Order Date";
                worksheet.Cells[2, 2].Value = order.Header.OrderDate;
                worksheet.Cells[3, 1].Value = "Order Currency";
                worksheet.Cells[3, 2].Value = order.Header.OrderCurrency;
                worksheet.Cells[4, 1].Value = "Expected Delivery Date";
                worksheet.Cells[4, 2].Value = order.Header.ExpectedDeliveryDate;

                // Parties
                worksheet.Cells[6, 1].Value = "Buyer ILN";
                worksheet.Cells[6, 2].Value = order.Parties.BuyerILN;
                worksheet.Cells[7, 1].Value = "Seller ILN";
                worksheet.Cells[7, 2].Value = order.Parties.SellerILN;
                worksheet.Cells[8, 1].Value = "Delivery Point ILN";
                worksheet.Cells[8, 2].Value = order.Parties.DeliveryPointILN;

                // Lines
                worksheet.Cells[10, 1].Value = "Lines";
                worksheet.Cells[11, 1].Value = "LineNumber";
                worksheet.Cells[11, 2].Value = "EAN";
                worksheet.Cells[11, 3].Value = "BuyerItemCode";
                worksheet.Cells[11, 4].Value = "ItemDescription";
                worksheet.Cells[11, 5].Value = "ItemType";
                worksheet.Cells[11, 6].Value = "OrderedQuantity";
                worksheet.Cells[11, 7].Value = "FreeOrderedQuantity";
                worksheet.Cells[11, 8].Value = "OrderedUnitPacksize";
                worksheet.Cells[11, 9].Value = "UnitOfMeasure";
                worksheet.Cells[11, 10].Value = "OrderedUnitNetPrice";
                worksheet.Cells[11, 11].Value = "ScheduledQuantity";

                int row = 12;
                foreach (var line in order.Lines)
                {
                    worksheet.Cells[row, 1].Value = line.LineNumber;
                    worksheet.Cells[row, 2].Value = line.EAN;
                    worksheet.Cells[row, 3].Value = line.BuyerItemCode;
                    worksheet.Cells[row, 4].Value = line.ItemDescription;
                    worksheet.Cells[row, 5].Value = line.ItemType;
                    worksheet.Cells[row, 6].Value = line.OrderedQuantity;
                    worksheet.Cells[row, 7].Value = line.FreeOrderedQuantity;
                    worksheet.Cells[row, 8].Value = line.OrderedUnitPacksize;
                    worksheet.Cells[row, 9].Value = line.UnitOfMeasure;
                    worksheet.Cells[row, 10].Value = line.OrderedUnitNetPrice;
                    worksheet.Cells[row, 11].Value = line.ScheduledQuantity;
                    row++;
                }

                package.SaveAs(new FileInfo(outputPath));
            }
        }
    }
}
