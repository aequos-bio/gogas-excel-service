using NUnit.Framework;
using ReportService.dto;
using ReportService.excel;

namespace ReportService.Test
{
    [TestFixture]
    public class ExcelProducerUnitTest
    {
        [SetUp]
        public void Setup()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        [Test]
        public void givenAnOrder_whenProducingExcel_thenExcelIsCorrect()
        {
            List<OrderProduct> productList = new List<OrderProduct> {
                new OrderProduct
                {
                    id = "p1",
                    name = "Prodotto 1",
                    boxWeight = 10,
                    unitOfMeasure = "KG",
                    unitPrice = new decimal(1.34)
                },
                new OrderProduct
                {
                    id = "p2",
                    name = "Prodotto 2",
                    boxWeight = 1,
                    unitOfMeasure = "PZ",
                    unitPrice = new decimal(1.1)
                },
                new OrderProduct
                {
                    id = "p3",
                    name = "Prodotto 3",
                    boxWeight = 12,
                    unitOfMeasure = "KG",
                    unitPrice = new decimal(3.57)
                },
                new OrderProduct
                {
                    id = "p4",
                    name = "Prodotto 4",
                    boxWeight = 5,
                    unitOfMeasure = "KG",
                    unitPrice = new decimal(0.90)
                }
            };

            List<User> usersList = new List<User> {
                mockUser(1, "Angela Giansiracusa"),
                mockUser(2, "Maria Ammairone"),
                mockUser(3, "Mariarita Rostirolla"),
                mockUser(4, "Flora Siracusa"),
                mockUser(5, "Roberta Pagetti")
            };

            List<OrderItem> totalOrder = new List<OrderItem> {
                mockOrderItem("p1", "u1", 1),
                mockOrderItem("p1", "u2", 1),
                mockOrderItem("p2", "u3", 1),
                mockOrderItem("p3", "u4", 1),
                mockOrderItem("p4", "u5", 1)
            };

            List<SupplierOrderItem> supplierOrder = new List<SupplierOrderItem> {
                mockSupplierOrderItem("p1", new decimal(1.34), 10, 1),
                mockSupplierOrderItem("p2", new decimal(1.1), 1, 1),
                mockSupplierOrderItem("p3", new decimal(3.57), 12, 1),
                mockSupplierOrderItem("p4", new decimal(0.90), 4, 1)
            };

            ExcelProducer excelProducer = new ExcelProducer();
            byte[] excel = excelProducer.exportOrder(productList, usersList, totalOrder, supplierOrder, false, false);
            File.WriteAllBytes("F:/dlorusso/report.xlsx", excel);
        }

        private User mockUser(int position, String name)
        {
            return new User
            {
                id = "u" + position,
                fullName = name,
                position = position,
                email = ""
            };
        }

        private OrderItem mockOrderItem(String productId, String userId, decimal quantity)
        {
            return new OrderItem
            {
                productId = productId,
                userId = userId,
                quantity = quantity
            };
        }

        private SupplierOrderItem mockSupplierOrderItem(String productId, decimal unitPrice, decimal boxWeight, decimal quantity)
        {
            return new SupplierOrderItem
            {
                productId = productId,
                unitPrice = unitPrice,
                boxWeight = boxWeight,
                quantity = quantity
            };
        }
    }                                                                      
}