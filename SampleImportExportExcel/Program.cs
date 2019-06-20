using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using SampleImportExportExcel.Messages;

namespace SampleImportExportExcel
{
    class Program
    {
        static string endpoint = "http://52.167.113.119/Pedido/{id}/entregue";
        static string token = "Bearer eyJhbGciOiJSUzI1NiIsImtpZCI6IkIyNUQ2MTFCRUVDNTREMDk3OEJBMEYwM0RFNUI0NjZFN0ZFMzhFRTQiLCJ0eXAiOiJKV1QiLCJ4NXQiOiJzbDFoRy03RlRRbDR1ZzhEM2x0R2JuX2pqdVEifQ.eyJuYmYiOjE1NjA5NzQwNjYsImV4cCI6MTg3NjMzNDA2NiwiaXNzIjoiaHR0cDovLzEwLjIzNC43Mi4xNjQiLCJhdWQiOlsiaHR0cDovLzEwLjIzNC43Mi4xNjQvcmVzb3VyY2VzIiwibWFya2V0cGxhY2VJbiJdLCJjbGllbnRfaWQiOiJtYXJrZXRwbGFjZUluIiwic3ViIjoiNWMxN2VmMDI1YWQ2ZDczMjc5MzI5NWFmIiwiYXV0aF90aW1lIjoxNTYwOTc0MDY2LCJpZHAiOiJsb2NhbCIsInJvbGUiOlsiTWFya2V0cGxhY2VJbi5BcGkuU2VsbGVyLkFkbWluIiwiTWFya2V0cGxhY2VJbi5Qb3J0YWxMb2ppc3RhLkFkbWluIl0sImlkIjoiNWMxN2VmMDI1YWQ2ZDczMjc5MzI5NWFmIiwiZW1haWwiOiJsaXZyb3MuY29tQGludGVncmFjYW8uY29tLmJyIiwidXNlcm5hbWUiOiJMaXZyb3MuY29tIiwiYXRpdm8iOiJ0cnVlIiwiaWRMb2ppc3RhIjoiMTAwNTkiLCJzY29wZSI6WyJlbWFpbCIsIm9wZW5pZCIsInByb2ZpbGUiLCJtYXJrZXRwbGFjZUluIiwib2ZmbGluZV9hY2Nlc3MiXSwiYW1yIjpbInB3ZCJdfQ.SUDFhNtH3pg1DoqdxHa0BPkOyYLQf2J1ENkf9kpx12sdR0MED3XhG-_QhJ7w8i5hbehPF1FQdYCUfO_kf6ov5ADbKTdXtv29sZSVJAW_v_ypUbnO9v13jqCts1W2AjmXs1uuq6IK_3YN-f6263PCHSQtddD_SmQtF0oGjK8fvzyE0nUCu3ZWIxk21KXiMtxh_vIZOnTpcqeGAhs1TllDZ8vf8zW7PVIBpd9Lrqg09UpOhqrNcKmqZdf2m1n7btor21400upaLhllQcUg5UYi8255OteLKjXFIEK3HS3yH0KAjy08op5dV-RnNjQ6NWf4gU50myGNOgrySNTybEX8juurHlHk4swBGwoqCSgM-MQwrhH2gq4LHM3d_9FFdByVoYPae4g4NSJk18BgTEFHetrS4CkKCkK4XOlSS-fy8Jz-ZxFAnVfkcrnxrjlEJmgdeiT7spbR57RbubsiOutcAKxJpBY8YjvI5EYoiDvNVBPqYEnpv1bUTWJAYKT5mfNSXYGUoSDJEKa-y7m-ECR7Hc76zG-HiEXazWzkbI9Ze_sHauUZsBbsosFXjj8W42vAU_fXrRf9IxuS48exGLKRNtwXIBDJ83pqWiPSsfIJIJfLrSTpI5Jg4XtxdG7tVjMzNp92ePDW0x5gdVhslS0T6XiwhMbdsmsub6EHWgq0gh8";

        static void Main(string[] args)
        {
            MainAsync(CancellationToken.None).Wait();
        }

        static async Task MainAsync(CancellationToken cancellationToken)
        {
            var file = GetFile();

            var importManager = new ImportManager();
            IEnumerable<Order> orders = importManager.ImportOrdersFromXlsx(file);

            foreach (var order in orders)
            {
                var message = OrderMessage.Create(order);

                var response = await SendPost(message, cancellationToken);
            }
        }

        static Stream GetFile()
        {
            string directory = Directory.GetCurrentDirectory();
            string filePath = Path.Combine(directory, "Files/entregas.xlsx");

            return File.OpenRead(filePath);
        }

        static async Task<HttpResponseMessage> SendPost(OrderMessage message, CancellationToken cancellationToken)
        {
            endpoint = endpoint.Replace("{id}", message.Numero);

            using (var client = new HttpClient())
            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            using (var httpContent = CreateHttpContent(message))
            {
                request.Content = httpContent;
                request.Headers.Add("Authorization", token);

                using (var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
                {
                    return response;
                }
            }
        }

        static HttpContent CreateHttpContent(object content)
        {
            HttpContent httpContent = null;

            if (content != null)
            {
                var ms = new MemoryStream();
                SerializeJsonIntoStream(content, ms);
                ms.Seek(0, SeekOrigin.Begin);
                httpContent = new StreamContent(ms);
                httpContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            }

            return httpContent;
        }

        static void SerializeJsonIntoStream(object value, Stream stream)
        {
            using (var sw = new StreamWriter(stream, new UTF8Encoding(false), 1024, true))
            using (var jtw = new JsonTextWriter(sw) { Formatting = Formatting.None })
            {
                var js = new JsonSerializer();
                js.Serialize(jtw, value);
                jtw.Flush();
            }
        }
    }
}
