using System;
using System.Threading.Tasks;
using UCClientWebService.Models;

namespace ConsoleAppTestUCClientWebService
{
    class Program
    {
        static void Main(string[] args)
        {

            ValidaAddin().Wait();
        }

        private static async Task ValidaAddin()
        {
            UCClientWebService.Services.AddInService addInService = new UCClientWebService.Services.AddInService();
            RequestID requestId = new RequestID()
            {
                id= "gwerZrM2"
            };

            var responseId = await addInService.IsValidAsync("https://www.ulmaconstruction.com/@@bim_form_api", requestId);
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
                Console.WriteLine("Addin invalid");
        }
    }
}
