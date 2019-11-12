using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UCClientWebService.Models;

namespace UCClientWebService.Services
{
    public class AddInService:ApiService
    {
        public ResponseID IsValidAsync(string urlEndPoint, RequestID request)
        {
            var response = this.Post<RequestID, ResponseID>(urlEndPoint, request);

            if (!response.IsSuccess)
            {
                return new ResponseID()
                {
                    id = request.id,
                    message = response.Message,
                    valid = false
                };
            }

            return (ResponseID)response.Result;

        }
    }
}
