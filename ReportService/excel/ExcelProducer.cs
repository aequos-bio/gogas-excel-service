using OfficeOpenXml;
using OfficeOpenXml.Style;
using ReportService.dto;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ReportService.excel
{
    public class ExcelProducer
    {
        private readonly String[] PRODUCTS_HEADER = { "Codice esterno", "Descrizione", "Codice fornitore", "Produttore", "Regione", "Categoria", "Unità di misura", "Peso collo", "Prezzo unitario", "Note", "Cadenza", "Acq. solo a collo", "Multiplo" };

        public byte[] exportOrder(List<OrderProduct> productList, List<User> usersList, List<OrderItem> totalOrder, List<SupplierOrderItem> supplierOrder, Boolean friends)
        {
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Dettaglio ordine");
            Font baseFont = new Font("Arial", 8, FontStyle.Regular);

            int row = 1, col = 1;

            row++;
            col = 6;

            foreach (User user in usersList)
            {
                ws.Cells[row, col].Value = (col - 5);
                ws.Cells[row, col].Style.Font.SetFromFont(new Font("Arial", 10, FontStyle.Bold));
                ws.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells[row + 1, col].Value = user.fullName;
                ws.Cells[row + 1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 1, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));

                if (user.role == User.FRIEND)
                    ws.Cells[row + 2, col].Value = "Amico di\r\n" + user.referralFullName;
                else
                    ws.Cells[row + 2, col].Value = user.phone + "\r\n" + user.email;

                ws.Cells[row + 2, col].Style.WrapText = true;
                ws.Cells[row + 2, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 2, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));

                col++;
            }

            ExcelRange range = ws.Cells[row + 1, 6, row + 2, col - 1];
            ApplyExcelBaseStyle(range, baseFont);

            row += 2;
            col = 1;

            ws.Cells[row, col++].Value = "Prodotto";

            ws.Cells[row, col].Value = "Prezzo\r\nper UM";
            ws.Cells[row, col++].Style.WrapText = true;

            ws.Cells[row, col++].Value = "UM";

            ws.Cells[row, col].Value = "Peso\r\nCollo";
            ws.Cells[row, col++].Style.WrapText = true;

            ws.Cells[row, col].Value = friends ? "Quantità\r\nritirata" : "Colli\r\nda ordinare";
            ws.Cells[row, col++].Style.WrapText = true;

            range = ws.Cells[row, 1, row, 5];
            ApplyExcelBaseStyle(range, baseFont);
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));

            row++;
            col = 1;

            //Scrivo righe
            foreach (var pr in productList)
            {
                SupplierOrderItem supplierOrderProduct = supplierOrder.Where(s => s.productId == pr.id).SingleOrDefault();

                var os = totalOrder.Where(g => g.productId == pr.id).OrderBy(g => g.userId);

                Decimal prezzoKg, pesoCollo, numeroColli;
                if (supplierOrderProduct != null)
                {
                    //Prodotto ordinato al produttore, prendo valori dall'ordine fornitore
                    prezzoKg = supplierOrderProduct.unitPrice;
                    pesoCollo = supplierOrderProduct.boxWeight;
                    numeroColli = supplierOrderProduct.quantity;
                }
                else
                {
                    //Prodotto non ordinato al produttore, prendo valori dal prodotto
                    prezzoKg = os.Any() ? os.First().unitPrice : pr.unitPrice;
                    pesoCollo = pr.boxWeight;
                    numeroColli = 0;
                }

                ws.Cells[row, col++].Value = pr.name;

                ws.Cells[row, col].Value = prezzoKg;
                ws.Cells[row, col++].Style.Numberformat.Format = "0.00 €";

                ws.Cells[row, col++].Value = pr.unitOfMeasure;

                ws.Cells[row, col].Value = pesoCollo;
                ws.Cells[row, col++].Style.Numberformat.Format = "0.00";

                ws.Cells[row, col].Value = numeroColli;
                ws.Cells[row, col].Style.Numberformat.Format = friends ? "0.000" : "0";
                ws.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, col++].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(235, 241, 222));

                foreach (var user in usersList)
                {
                    var dt = os.Where(g => g.userId == user.id).FirstOrDefault();
                    if (dt != null)
                        //Se esporto il file per ripartizione amici riporto le qta ordinate altrimenti è per riepilogo ordine e quindi riporto qta ritirata
                        ws.Cells[row, col].Value = dt.quantity;
                    col++;
                }

                row++;
                col = 1;
            }

            col = 5;

            ws.Cells[row, col++].Value = "Totali";
            foreach (var itemU in usersList)
            {
                ws.Cells[row, col].Formula = "sumproduct(" + ws.Cells[5, 2].Address + ":" + ws.Cells[row - 1, 2].Address + "," + ws.Cells[5, col].Address + ":" + ws.Cells[row - 1, col].Address + ")";
                col++;
            }
            ws.Cells[row, 6, row, col].Style.Numberformat.Format = "0.00 €";

            ws.Cells[5, 6, row - 1, 5 + usersList.Count()].Style.Numberformat.Format = "0.000";

            range = ws.Cells[5, 1, row - 1, 5 + usersList.Count()];
            ApplyExcelBaseStyle(range, baseFont);

            ws.Cells[5, 1, row - 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            range = ws.Cells[row, 5, row, 5 + usersList.Count()];
            ApplyExcelBaseStyle(range, baseFont);
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));

            ws.Cells[1, 1, row, 5 + usersList.Count()].AutoFitColumns();

            ws.View.FreezePanes(5, 6);

            return pck.GetAsByteArray();
        }

        public byte[] exportProducts(IList<PriceListProduct> products)
        {
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Prodotti");

            Font baseFont = new Font("Arial", 8, FontStyle.Regular);

            //Creating data
            List<Object[]> excelData = new List<Object[]> { PRODUCTS_HEADER };
            excelData.AddRange(products.Select(p => new Object[] {
                                                        p.externalId,
                                                        p.name,
                                                        p.supplierExternalId,
                                                        p.supplierName,
                                                        p.supplierProvince,
                                                        p.category,
                                                        p.unitOfMeasure,
                                                        p.boxWeight,
                                                        p.unitPrice,
                                                        p.notes,
                                                        p.frequency,
                                                        p.wholeBoxesOnly ? "S" : "N",
                                                        p.multiple
                                                    }));

            ws.Cells[1, 1].LoadFromArrays(excelData);

            //Formatting Header
            ExcelRange range = ws.Cells[1, 1, 1, PRODUCTS_HEADER.Count()];
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));
            range.Style.Font.Color.SetColor(Color.White);
            range.Style.Font.SetFromFont(new Font("Arial", 8, FontStyle.Bold));

            //Formatting font
            ws.Cells[2, 1, excelData.Count(), PRODUCTS_HEADER.Count()].Style.Font.SetFromFont(baseFont);

            //Formatting number fields
            ws.Cells[2, 8, excelData.Count(), 8].Style.Numberformat.Format = "0.00";
            ws.Cells[2, 9, excelData.Count(), 9].Style.Numberformat.Format = "0.00 €";

            for (int i = 1; i <= PRODUCTS_HEADER.Count(); i++)
                ws.Column(i).AutoFit();

            //TODO: fare secondo foglio con legenda e spiegazioni

            return pck.GetAsByteArray();
        }

        private void ApplyExcelBaseStyle(ExcelRange range, Font baseFont)
        {
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 128, 0));
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Color.SetColor(Color.FromArgb(0, 128, 0));
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Color.SetColor(Color.FromArgb(0, 128, 0));
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 128, 0));

            range.Style.Font.SetFromFont(baseFont);

            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
    }
}
