using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UCClientWebService.Models;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            ValidaAddin();           
        }

        private static  void ValidaAddin()
        {
            UCClientWebService.Services.AddInService addInService = new UCClientWebService.Services.AddInService();
            RequestID requestId = new RequestID()
            {
                id = "WXdgcIZ"
            };

            var responseId = addInService.IsValidAsync("https://www.ulmaconstruction.com/@@bim_form_api", requestId);
            if (responseId != null)
            {
                if (responseId.valid)
                    Console.WriteLine("Addin valid");
                else
                {
                    Console.WriteLine("Addin invalid");
                    Console.WriteLine(responseId.message); //Mensaje personalizado enviado desde el servidor
                }

            }
            else
                Console.WriteLine("Connection error");
        }
    }
}
