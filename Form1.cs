using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OfficeOpenXml;
using System.Text;
using System.Security.Cryptography;
using Microsoft.Win32;

namespace TesteParalelo_2025
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        public void Iniciar()
        {
            try
            {
                List<Aluno> registros = LerDadosDoExcel(@"C:\Users\romenique.santos\Downloads\COP\Pasta1.xlsx");
                int totalRegistros = registros.Count;
                List<Task> tarefas = new List<Task>();

                int numeroDeProcessos = 3;

                // Calcula quantos registros cada processo (janela) deve processar.
                // Usamos Math.Ceiling para garantir que, se houver um número ímpar de registros,
                // o último processo ainda receba todos os registros restantes.
                int registrosPorJanela = (int)Math.Ceiling((double)totalRegistros / numeroDeProcessos);

                // Inicia um loop que irá criar os processos para cada janela.
                for (int i = 0; i < numeroDeProcessos; i++)
                {
                    // Calcula o índice inicial na lista de registros para este processo específico.
                    // Por exemplo, se i for 0, o índice inicial será 0 (primeiro grupo de registros).
                    // Se i for 1, o índice inicial será igual a 'registrosPorJanela' (segundo grupo).
                    int indiceInicial = i * registrosPorJanela;

                    // Calcula quantos registros este processo deve pegar.
                    // Usamos Math.Min para garantir que não tentemos acessar mais registros do que existem na lista.
                    int quantidadeRegistros = Math.Min(registrosPorJanela, totalRegistros - indiceInicial);

                    // Verifica se há registros a serem processados. Se a quantidade for maior que 0,
                    // significa que existem registros disponíveis para este processo.
                    if (quantidadeRegistros > 0)
                    {
                        // Pega uma parte dos registros da lista original usando GetRange,
                        // começando no índice inicial e pegando a quantidade calculada.
                        var parteRegistros = registros.GetRange(indiceInicial, quantidadeRegistros);


                        // Adiciona uma nova tarefa à lista de tarefas. Cada tarefa executará o método
                        // CadastrarRegistros passando os registros para processar e o ID da janela.
                        int id = i + 1;
                        tarefas.Add(Task.Run(() => CadastrarRegistros(parteRegistros, id)));
                    }
                }

                // Espera até que todas as tarefas na lista sejam concluídas antes de continuar.
                Task.WaitAll(tarefas.ToArray());
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void CadastrarRegistros(List<Aluno> registros, int windowId)
        {
            IWebDriver driver = null;

            try
            {
                driver = IniciarWebDriver(driver);
                driver.Navigate().GoToUrl("https://www.google.com/");
                Thread.Sleep(5000);
                foreach (var registro in registros)
                {
                    string CPF = registro.CPF;
                    try
                    {
                        Pesquisar(driver, CPF);
                        Thread.Sleep(2000);

                        // Supondo que o registro foi bem-sucedido
                        log(windowId, CPF, "Sucesso");
                    }
                    catch (Exception ex)
                    {
                        // Adiciona erro ao resultado
                        log(windowId, CPF, "Erro " + ex.Message);
                        EncerrarDriver(driver);
                    }
                }

                EncerrarDriver(driver);
            }
            catch (Exception ex)
            {
                // Em caso de erro geral no navegador
                log(windowId, "Geral ", "Erro " + ex.Message);
                EncerrarDriver(driver);
            }
        }

        public void EncerrarDriver(IWebDriver driver)
        {
            do
            {
                try
                {
                    driver.Close();
                    driver.Quit();
                    driver.Dispose();
                }
                catch (Exception)
                {
                    break;
                }
            } while (true);
        }

        public IWebDriver IniciarWebDriver(IWebDriver driver_)
        {
            do
            {
                EncerrarDriver(driver_);
                try
                {
                    //Chrome Option
                    var chromeOptions = new ChromeOptions();

                    //ChromeDriverService service = ChromeDriverService.CreateDefaultService(@"C:\Roma Automações\Driver\");
                    ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                    service.SuppressInitialDiagnosticInformation = true;
                    service.HideCommandPromptWindow = true;

                    //Iniciar o Chrome
                    IWebDriver driver = new ChromeDriver(service, chromeOptions);

                    return driver;
                }
                catch (Exception)
                {

                }
            } while (true);
        }

        public void Pesquisar(IWebDriver driver, string CPF)
        {
            driver.FindElement(By.Id("APjFqb")).Clear();
            driver.FindElement(By.Id("APjFqb")).SendKeys(CPF);
        }

        public void log(int windowId, string CPF, string MSG)
        {
            // Caminho do arquivo CSV para a janela atual
            string csvFilePath = $@"C:\Users\romenique.santos\Downloads\COP\resultado_janela_{windowId}.csv";

            bool cabecalho;
            if (File.Exists(csvFilePath))
                cabecalho = true;
            else
                cabecalho = false;

            StreamWriter writer = new StreamWriter(csvFilePath, true, Encoding.Default);

            if (!cabecalho)
                writer.WriteLine("CPF;Mensagem"); // Escreve o cabeçalho

            string linhas = CPF + ";" + MSG;
            writer.WriteLine(linhas);
            Console.WriteLine(linhas);

            writer.Close();
            writer.Dispose();
        }

        static List<Aluno> LerDadosDoExcel(string caminho)
        {
            var registros = new List<Aluno>();

            // Configurar o EPPlus para usar a leitura de arquivos
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(caminho)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Pega a primeira planilha
                int totalRows = worksheet.Dimension.Rows;

                for (int row = 2; row <= totalRows; row++) // Começa na linha 2 para ignorar cabeçalho
                {
                    registros.Add(new Aluno()
                    {
                        CPF = worksheet.Cells[row, 1].Text, // Coluna A
                        NOME = worksheet.Cells[row, 2].Text, // Coluna B
                        DIVIDA = worksheet.Cells[row, 3].Text, // Coluna C
                    });
                }
            }

            return registros;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Iniciar();
        }
    }
}


public class Aluno
{
    public string CPF { get; set; }
    public string NOME { get; set; }
    public string DIVIDA { get; set; }
}