using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;
using System.Diagnostics;

namespace BotWhatsapp
{
    class RPACorreios
    {
        public static object CsharpToExcel { get; private set; }

        static void Main(string[] args)
        {
            run();


            static void run()
            {
                capturaInfoEnderecos();
            
            }

            static void capturaInfoEnderecos()
            {
                ChromeDriver driver = new ChromeDriver();
                
                string url = "https://buscacepinter.correios.com.br/app/endereco/index.php";

                driver.Navigate().GoToUrl(url);
                driver.Manage().Window.Maximize();

                List<String> cepList = new List<String>()
                {
                    "17033470"
                };
                
                foreach (var cep in cepList)
                {
                    var inputCep = driver.FindElement(OpenQA.Selenium.By.Id("endereco"));
                    inputCep.SendKeys(cep);

                    var btnBuscar = driver.FindElement(OpenQA.Selenium.By.Id("btn_pesquisar"));
                    btnBuscar.Click();

                    int tentativas = 0;

                    Thread.Sleep(1000);

                    while (tentativas <= 3)
                    {
                        try
                        {
                           var logradouro = driver.FindElement(OpenQA.Selenium.By.XPath("//*[@id='resultado-DNEC\']/tbody/tr/td[1]"));

                            var bairro = driver.FindElement(OpenQA.Selenium.By.XPath("//*[@id=\"resultado-DNEC\"]/tbody/tr/td[2]"));
                            var localidade = driver.FindElement(OpenQA.Selenium.By.XPath("//*[@id=\"resultado-DNEC\"]/tbody/tr/td[3]"));

                            geraPlanilha(logradouro.Text, bairro.Text, localidade.Text, cep);
                        }
                        catch(Exception e)
                        {
                            Thread.Sleep(1000);
                            tentativas++;
                        }
                    }
                }
            }

            static void geraPlanilha(string logradouro, string bairro, string localidade, string cep)
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Enderecos");
                    worksheet.Cell("A1").Value = "Logradouro/Nome";
                    worksheet.Cell("B1").Value = "Bairro/Distrito";
                    worksheet.Cell("C1").Value = "Localidade/UF";
                    worksheet.Cell("D1").Value = "CEP";

                    worksheet.Cell("A2").Value = logradouro;
                    worksheet.Cell("B2").Value = bairro;
                    worksheet.Cell("C2").Value = localidade;
                    worksheet.Cell("D2").Value = cep;

                    workbook.SaveAs(@"c:\temp\planilhaEnderecos.xlsx");
                    Process.Start(new ProcessStartInfo(@"c:\temp\planilhaEnderecos.xlsx") { UseShellExecute = true });
                }
            }

        }
    }
}