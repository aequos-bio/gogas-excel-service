using NUnit.Framework;
using ReportService.excel;

namespace ReportService.Test
{
    [TestFixture]
    public class ExcelReaderUnitTest
    {
        [SetUp]
        public void Setup()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        [Test]
        public void givenACorrectExcelPriceList_whenReadingProducts_thenAllProductsAreRead()
        {
            byte[] priceListContent = File.ReadAllBytes("resources/pricelist_ok.xlsx");
            ExcelReader excelReader = new ExcelReader();
            List<dto.PriceListProduct> priceListProducts = excelReader.extractProducts(priceListContent, "xlsx");
            Assert.That(priceListProducts.Count, Is.EqualTo(91));
        }

        [Test]
        public void givenAnExcelPriceListWithNumericSupplierIdAndEmptyRows_whenReadingProducts_thenAllProductsAreRead()
        {
            byte[] priceListContent = File.ReadAllBytes("resources/pricelist_empty_lines.xlsx");
            ExcelReader excelReader = new ExcelReader();
            List<dto.PriceListProduct> priceListProducts = excelReader.extractProducts(priceListContent, "xlsx");
            Assert.That(priceListProducts.Count, Is.EqualTo(91));
        }
    }
}