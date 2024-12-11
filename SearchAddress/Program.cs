using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Text;
using System.IO;
using Dto.Api;
using System;

namespace SearchAdress
{
    class Program
    {
        static string RemoveAccents(string text)
        {
            string result = "";

            if (text.Length > 0)
            {
                try
                {
                    var Encode8Bits = Encoding.GetEncoding(1251).GetBytes(text);
                    var String7Bits = Encoding.ASCII.GetString(Encode8Bits);
                    var RemoveAccent = new Regex("[^a-zA-Z0-9]=-_/");
                    result = RemoveAccent.Replace(String7Bits, " ");
                }
                catch (Exception)
                {
                    return result;
                }
            }

            return result;
        }
 
        static async Task<Data> CheckAddress(string zipcode)
        {
            try
            {
                string baseURL = "https://viacep.com.br";

                var client = new HttpClient();
                client.Timeout = TimeSpan.FromHours(4);
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json");

                var response = new HttpResponseMessage();

                response = client.GetAsync($"{baseURL}/ws/" + $"{zipcode}" + "/json").Result;
                var resposta = JsonConvert.DeserializeObject<Data>(RemoveAccents(await response.Content.ReadAsStringAsync()));

                return resposta;
            }
            catch (Exception)
            {
                return null;
            }
        }

        static void Main(string[] args)
        {
            string[] zipcode = new string[] { "77060144" , "60341050", "29166048", "69908734", "78556674", "76801098", "88655970" };

            using (var excelPackage = new ExcelPackage())
            {
                excelPackage.Workbook.Properties.Author = "Address Data";
                excelPackage.Workbook.Properties.Title = "Excel";

                var sheet = excelPackage.Workbook.Worksheets.Add("Spreadsheet 1");
                sheet.Name = "Spreadsheet 1";

                var count = 1;
                string[] exceltitles = new String[] { "ZIP CODE", "PUBLIC PLACE", "COMPLEMENT", "NEIGHBORHOOD", "UF", "STATE" };
                foreach (string word in exceltitles)
                {
                    sheet.Cells[1, count++].Value = word;
                }

                var dataexcel = new String[7];
                int position = 2;
                for (int i = 0; i < dataexcel.Length; i++)
                {
                    var data = CheckAddress(zipcode[i]).Result;

                    if (data != null)
                    {
                        dataexcel[0] = data.cep;
                        dataexcel[1] = data.logradouro.ToUpper();
                        dataexcel[2] = string.IsNullOrEmpty(data.complemento) ? "-" : data.complemento.ToUpper();
                        dataexcel[3] = data.bairro.ToUpper();
                        dataexcel[4] = data.uf.ToUpper();
                        dataexcel[5] = data.estado.ToUpper();
                        dataexcel[6] = data.ddd;
                    }
                    else
                    {
                        dataexcel[0] = zipcode[i];
                        dataexcel[1] = "-";
                        dataexcel[2] = "-";
                        dataexcel[3] = "-";
                        dataexcel[4] = "-";
                        dataexcel[5] = "-";
                        dataexcel[6] = "-";
                    }

                    //SAVE IN EXCEL
                    int countResult = 1;
                    foreach (var value in dataexcel)
                    {
                        sheet.Cells[position, countResult++].Value = value;                       
                    }
                    position++;
                    //CHECK YOUR COMPUTER PATH
                    string path = @"C:\Users\PC\Documents\Path\Addresses.xlsx";
                    File.WriteAllBytes(path, excelPackage.GetAsByteArray());
                }

                System.Environment.Exit(1);
            }

        }
    }
}