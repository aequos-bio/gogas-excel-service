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

        public byte[] exportOrder(List<OrderProduct> productList, List<User> usersList, List<OrderItem> totalOrder, 
                                  List<SupplierOrderItem> supplierOrder, Boolean friends, Boolean addWeightColumns)
        {
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Dettaglio ordine");
            Font baseFont = new Font("Arial", 8, FontStyle.Regular);

            int row = 1, col = 1;
            int fixedCols = addWeightColumns ? 8 : 6;

            row++;
            col = fixedCols;

            foreach (User user in usersList)
            {
                ws.Cells[row, col].Value = user.position;
                ws.Cells[row, col].Style.Font.SetFromFont(new Font("Arial", 10, FontStyle.Bold));
                ws.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells[row + 1, col].Value = user.fullName;
                ws.Cells[row + 1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 1, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));

                if (addWeightColumns)
                    ws.Cells[row + 2, col].Value = "Ordinato";
                else if (user.role == User.FRIEND)
                    ws.Cells[row + 2, col].Value = "Amico di\r\n" + user.referralFullName;
                else
                    ws.Cells[row + 2, col].Value = user.phone + "\r\n" + user.email;

                ws.Cells[row + 2, col].Style.WrapText = true;
                ws.Cells[row + 2, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 2, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));

                if (addWeightColumns)
                {
                    ws.Cells[row, col, row, col + 2].Merge = true;
                    ws.Cells[row + 1, col, row + 1, col + 2].Merge = true;
                    ws.Cells[row + 2, ++col].Value = "Peso";
                    ws.Cells[row + 2, ++col].Value = "Valore €";
                }

                col++;
            }

            ExcelRange range = ws.Cells[row + 1, fixedCols, row + 2, col - 1];
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

            if (addWeightColumns)
            {
                ws.Cells[row, col].Value = "Peso\r\ntotale";
                ws.Cells[row, col++].Style.WrapText = true;

                ws.Cells[row, col].Value = "Costo\r\ntotale";
                ws.Cells[row, col++].Style.WrapText = true;
            }

            range = ws.Cells[row, 1, row, col-1];
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

                if (addWeightColumns)
                {
                    ws.Cells[row, col].Formula = buildSumFormula(usersList.Count, row, col, 3, ws);
                    ws.Cells[row, col++].Style.Numberformat.Format = "0.000";

                    ws.Cells[row, col].Formula = buildSumFormula(usersList.Count, row, col, 3, ws); ;
                    ws.Cells[row, col++].Style.Numberformat.Format = "0.00 €";
                }

                foreach (var user in usersList)
                {
                    var dt = os.Where(g => g.userId == user.id).FirstOrDefault();
                    if (dt != null)
                        //Se esporto il file per ripartizione amici riporto le qta ordinate altrimenti è per riepilogo ordine e quindi riporto qta ritirata
                        ws.Cells[row, col].Value = dt.quantity;

                    if (addWeightColumns)
                    {
                        col += 2;
                        ws.Cells[row, col].Formula = "if(isblank(" + ws.Cells[row, col - 1].Address + "),\"\",product(" + ws.Cells[row, 2].Address + "," + ws.Cells[row, col - 1].Address + "))";
                    }

                    col++;
                }

                row++;
                col = 1;
            }

            col = fixedCols;

            ws.Cells[row, col - 1].Value = "Totali";
            foreach (var itemU in usersList)
            {
                if (addWeightColumns)
                {
                    col += 2;
                    ws.Cells[row, col].Formula = "sum(" + ws.Cells[5, col].Address + ":" + ws.Cells[row - 1, col].Address + ")";
                }
                else
                {
                    ws.Cells[row, col].Formula = "sumproduct(" + ws.Cells[5, 2].Address + ":" + ws.Cells[row - 1, 2].Address + "," + ws.Cells[5, col].Address + ":" + ws.Cells[row - 1, col].Address + ")";
                }
                col++;
            }
            int userColumns = usersList.Count() * (addWeightColumns ? 3 : 1);

            ws.Cells[row, fixedCols, row, (fixedCols - 1) + userColumns].Style.Numberformat.Format = "0.00 €";
            ws.Cells[5, fixedCols, row - 1, (fixedCols - 1) + userColumns].Style.Numberformat.Format = "0.000";

            range = ws.Cells[5, 1, row - 1, (fixedCols - 1) + userColumns];
            ApplyExcelBaseStyle(range, baseFont);

            ws.Cells[5, 1, row - 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            range = ws.Cells[row, (fixedCols - 1), row, (fixedCols - 1) + userColumns];
            ApplyExcelBaseStyle(range, baseFont);
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));

            col = fixedCols;
            if (addWeightColumns)
            {
                //formato colore differente per le colonne aggiuntive sul peso
                foreach (var itemU in usersList)
                {
                    range = ws.Cells[4, col + 1, row, col + 1];
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 241, 196));

                    range = ws.Cells[4, col + 2, row, col + 2];
                    range.Style.Numberformat.Format = "0.00 €";
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 199, 201));

                    col += 3;
                }
            }

            ws.Cells[1, 1, row, (fixedCols - 1) + userColumns].AutoFitColumns();

            ws.View.FreezePanes(5, fixedCols);

            return pck.GetAsByteArray();
        }

        private String buildSumFormula(int times, int row, int col, int interval, ExcelWorksheet ws)
        {
            string formula = "";
            for (int i = 1; i <= times; i++)
                formula += ws.Cells[row, col + (3 * i)].Address + ",";
            formula = formula.Substring(0, formula.Length - 1);
            return "sum(" + formula + ")";
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
