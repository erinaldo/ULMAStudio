using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using UCClientWebService.Models;

namespace UCClientWebService.Services
{
    public class ApiService
    {
        
        public Response Post<T, M>( string urlEndPoint, T model)
        {
            try
            {

                var requestString = JsonConvert.SerializeObject(model);
                var content = new StringContent(requestString, Encoding.UTF8, "application/json");
                var client = new HttpClient();
                
                var response = client.PostAsync(urlEndPoint, content).Result;
               
                var answer = response.Content.ReadAsStringAsync().Result;
                if (!response.IsSuccessStatusCode)
                {
                    return new Response
                    {
                        IsSuccess = false,
                        Message = ((int)response.StatusCode).ToString() + " - " +  response.ReasonPhrase
                    };
                }

                var obj = JsonConvert.DeserializeObject<M>(answer);
                return new Response
                {
                    IsSuccess = true,
                    Result = obj,
                };
            }
            catch (Exception ex)
            {
                
                return new Response
                {
                    IsSuccess = false,
                    Message = ex.InnerException.ToString(),
                };
            }
        }

    }
}
