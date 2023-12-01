using ExcelDataReader;
using ReportService.dto;
using ReportService.exception;
using System;
using System.Collections.Generic;
using System.IO;

namespace ReportService.excel
{
    public class ExcelReader
    {
        public List<PriceListProduct> extractProducts(byte[] excelContent, String excelType)
        {
            Int32 rowIndex = -1, colIndex = -1;

            if (excelContent == null || excelContent.Length == 0)
                throw new Exception("Invalid content");

            try
            {
                using (Stream excelStream = new MemoryStream(excelContent))
                {
                    using (IExcelDataReader excelReader = initReader(excelType, excelStream))
                    {

                        List<PriceListProduct> products = new List<PriceListProduct>();

                        //init row and col indexes to scan excel file
                        rowIndex = 1;
                        colIndex = 0;

                        //Starting from second row, skipping header
                        excelReader.Read();

                        //Data Reader methods
                        while (excelReader.Read())
                        {

                            rowIndex++;

                            colIndex = 0;

                            if (excelReader.GetValue(colIndex) == null)
                            {
                                break;
                            }

                            PriceListProduct productDictionary = new PriceListProduct();
                            productDictionary.externalId = excelReader.GetValue(colIndex++).ToString();
                            productDictionary.name = excelReader.GetString(colIndex++);
                            productDictionary.supplierExternalId = excelReader.GetValue(colIndex++).ToString();
                            productDictionary.supplierName = excelReader.GetString(colIndex++);
                            productDictionary.supplierProvince = excelReader.GetString(colIndex++);
                            productDictionary.category = excelReader.GetString(colIndex++);
                            productDictionary.unitOfMeasure = excelReader.GetString(colIndex++);

                            decimal boxWeight = Convert.ToDecimal(excelReader.GetDouble(colIndex++));
                            if (boxWeight <= 0)
                                throw new Exception("Il peso collo deve essere maggiore di zero");
                            productDictionary.boxWeight = boxWeight;

                            productDictionary.unitPrice = Convert.ToDecimal(excelReader.GetDouble(colIndex++));
                            productDictionary.notes = excelReader.FieldCount > 9 ? excelReader.GetString(colIndex++) : null;
                            productDictionary.frequency = excelReader.FieldCount > 10 ? excelReader.GetString(colIndex++) : null;
                            productDictionary.wholeBoxesOnly = excelReader.FieldCount > 11 ? excelReader.GetString(colIndex++).ToUpper() == "S" : false;

                            if (excelReader.FieldCount > 12 && excelReader.GetValue(colIndex) != null)
                                productDictionary.multiple = Convert.ToDecimal(excelReader.GetDouble(colIndex++));

                            //controllo per evitare che il flag "solo collo" sia attivo con unità di misura collo vuota (che succede quando il peso collo è 1), che fa andare in errore l'app
                            if (productDictionary.wholeBoxesOnly && productDictionary.boxWeight <= 1)
                                productDictionary.wholeBoxesOnly = false;

                            products.Add(productDictionary);
                        }

                        return products;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ExcelExtractionException(ex, rowIndex, colIndex);
            }
        }

        private static IExcelDataReader initReader(string excelType, Stream excelStream)
        {
            IExcelDataReader excelReader;
            if (excelType == "xls")
                //Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(excelStream);
            else if (excelType == "xlsx")
                //Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelStream);
            else
                throw new Exception("Invalid format");
            return excelReader;
        }
    }
}
