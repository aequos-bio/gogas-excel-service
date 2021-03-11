using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ReportService.dto;
using ReportService.excel;
using ReportService.exception;
using ReportService.extension;

namespace ReportService.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        // GET api/values
        [HttpPost]
        [Route("order")]
        public IActionResult Order([FromBody] OrderExportRequest request)
        {
            ExcelProducer producer = new ExcelProducer();
            byte[] excelContent = producer.exportOrder(request.products, request.users, request.userOrder, request.supplierOrder, request.friends, request.addWeightColumns);
            return new FileContentResult(excelContent, "application/octet-stream");
        }

        [HttpPost]
        [Route("products/generate")]
        public IActionResult GenerateProductList([FromBody] List<PriceListProduct> products)
        {
            ExcelProducer producer = new ExcelProducer();
            byte[] excelContent = producer.exportProducts(products);
            return new FileContentResult(excelContent, "application/octet-stream");
        }

        [HttpPost]
        [Route("products/extract/{excelType}")]
        [Consumes("application/octet-stream")]
        [Produces("application/json")]
        public async Task<IActionResult> RawBinaryDataManual(String excelType)
        {
            try
            {
                byte[] excelContent = await Request.GetRawBodyBytesAsync();
                ExcelReader reader = new ExcelReader();
                List<PriceListProduct> products = reader.extractProducts(excelContent, excelType);
                return new JsonResult(new {
                    priceListItems = products
                });
            }
            catch (ExcelExtractionException extractionEx)
            {
                return new JsonResult(new {
                    error = new ExtractionError(extractionEx)
                });
            }
        }
    }
}
