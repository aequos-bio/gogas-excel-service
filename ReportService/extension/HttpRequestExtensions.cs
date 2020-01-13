using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ReportService.extension
{
    public static class HttpRequestExtensions
    {
        /// <summary>
        /// Retrieves the raw body as a byte array from the Request.Body stream
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        public static async Task<byte[]> GetRawBodyBytesAsync(this HttpRequest request)
        {
            using (var stream = new MemoryStream(2048))
            {
                await request.Body.CopyToAsync(stream);
                return stream.ToArray();
            }
        }
    }
}
