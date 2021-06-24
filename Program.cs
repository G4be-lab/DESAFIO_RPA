using System;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;

namespace CEPFacil
{

    ////////////////////////////////////////////////////////////
    //                                                        //
    //                                                        //
    //                Done by Gabriel Martins                 //
    //                  Date of: 24/06/2021                   //
    //                                                        //
    //              CEP Searcher made using Selenium          //
    //               with support for Multi-threading         //    
    //                                                        //
    ////////////////////////////////////////////////////////////
   
    class Program
    {
        #region Declarations
        static string root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\\..\\"));
        public static string inputPath = root + @"\inputFile.xlsx";
        public static string outputPath = root + @"\OutputFile.xlsx";
        public static string driverFile = root;
        public static string url = @"https://buscacepinter.correios.com.br/app/endereco/index.php";
        public static Thread[] threads;
        public static Worker[] workers;
        public static XLWorkbook inputWb;
        public static XLWorkbook outputWb;
        public static int lineCount, linesFound;
        static int indexCount;
        public static bool once = false;
        #endregion

        //# - Params
        public static int threadsToUse = 8;
        public static bool useHeadless = true;
        //#

        #region Static_Functions
        public static void Write(Data data)
        {
            outputWb.Worksheet(1).Cell(indexCount+2, 1).Value = data.cep;
            outputWb.Worksheet(1).Cell(indexCount+2, 2).Value = data.nome;
            outputWb.Worksheet(1).Cell(indexCount+2, 3).Value = data.bairro;
            outputWb.Worksheet(1).Cell(indexCount+2, 4).Value = data.estado;
            outputWb.Worksheet(1).Cell(indexCount + 2, 5).Value = data.timestamp.ToString();
            indexCount++;
            outputWb.Save();
        }

        static int[] GetCEPRange()
        {        
            var n = inputWb.Worksheet(1).Cell(lineCount + 2, 1).Value.ToString();
            lineCount++;
            {
                if (lineCount == 1)
                {
                    Console.Clear();
                }
            }

            n = n.Replace("-", "");            
            n = n.Replace(" a ", " ");
            
            return new int[] { 
                int.Parse(n.Split(' ')[0]),
                int.Parse(n.Split(' ')[1]) 
            };
        }
        static void CloseBrowsers(object sender, EventArgs e)
        {
            for (int i = 0; i < threadsToUse; i++)
                workers[i].browser.Dispose();            
        }
#endregion
        
        static void Main(string[] args)
        {            
            //Set-up Excel files
            inputWb = new XLWorkbook(inputPath);
            outputWb = new XLWorkbook();
            outputWb.AddWorksheet("CEPs Information");
            outputWb.SaveAs(outputPath);

            AppDomain.CurrentDomain.ProcessExit += new EventHandler(CloseBrowsers); // close browsers on exit

            indexCount = 0;
            lineCount = 0;
            linesFound = 1;
           
            while (!string.IsNullOrEmpty((string)inputWb.Worksheet(1).Cell(linesFound+1, 1).Value)) // Reads first column down until blank space, ignores header
                linesFound++;            

            for (int i = 1; i <= 5; i++) 
                outputWb.Worksheet(1).Cell(1, i).Value = ((Fields)i).ToString("g").Replace("_"," "); //Header

            Console.Clear();
            Console.WriteLine(linesFound + " Lines found.");

            //Set-up workers and threads
            threads = new Thread[threadsToUse];
            workers = new Worker[threadsToUse];
          
            for (int i = 0; i < threadsToUse; i++)
            {
                workers[i] = new Worker(i);
                threads[i] = new Thread(new ThreadStart(workers[i].Work));
                threads[i].Name = "Thread " + i.ToString();
            }
            
            
            //While has stuff to read, do so.
            bool hasStuff = true;
            while (hasStuff)
            {
                
                for (int i = 0; i < threadsToUse; i++)
                {
                    if (lineCount < linesFound && !workers[i].busy)
                    {                        
                        var ceps = GetCEPRange();
                        workers[i].UpdateCEPs(ceps[0], ceps[1]);
                        threads[i] = new Thread(new ThreadStart(workers[i].Work));
                        threads[i].Name = "Thread " + i.ToString();
                        threads[i].Start();                        
                    }
                    hasStuff = (lineCount + 1) < linesFound;
                }                
            }
            Console.ReadLine();
        }
        
    }

    public class Worker
    {
        public ChromeDriver browser=null;
        public int cepStart;
        public int cepEnd;
        public int id;
        public bool busy = false;
        public Worker(int id) { this.id = id; }
            
        public void UpdateCEPs(int a, int b)
        {
            cepStart = a;
            cepEnd = b;
        }
        public void Work()
        {
            busy = true;
            List<Data> dataList = new List<Data>();

            var browserOptions = new ChromeOptions();
            if (Program.useHeadless)
            {
                browserOptions.AddArgument("headless");
            }
            browserOptions.AddArgument("--window-size=1280,720");

            if (browser == null)
                browser = new ChromeDriver(Program.driverFile, browserOptions);

            browser.Manage().Window.Minimize();

            int attempts=0;
            while (attempts < 2)
            {
                try
                {
                    browser.Navigate().GoToUrl(Program.url);
                    break;
                }
                catch (WebDriverException e) { }
                attempts++;
            }

            browser.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            browser.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);

            WebDriverWait wait = new WebDriverWait(browser, TimeSpan.FromSeconds(60));

            bool outLoop = false;
            int CEPcount = 0;
            while (!outLoop)
            {
                try
                {
                    int actualCep = cepStart + CEPcount;
                    CEPcount++;

                    var form = browser.FindElementById("endereco");
                    form.Click();
                    form.SendKeys((actualCep).ToString()); // CEP value inserted


                    browser.FindElementById("btn_pesquisar").Click(); //Submit

                    wait.Until(bwr => bwr.FindElement(By.Id("btn_voltar")));

                    if (!Program.once) // Clean the console
                    {
                        Program.once = true;
                        Console.Clear();
                    }

                    //Shorten timeout to optimize
                    browser.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(.40); // higher than .3 to prevent skip
                    bool hasResults = browser.FindElementsByXPath("//div[@id='mensagem-resultado']/h4").Count > 0;
                    browser.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

                    if (hasResults)
                    {
                        string cep = browser.FindElementByXPath("//table/tbody/tr/td[4]").Text;
                        string name = browser.FindElementByXPath("//table/tbody/tr/td[1]").Text;
                        string district = browser.FindElementByXPath("//table/tbody/tr/td[2]").Text;
                        string state = browser.FindElementByXPath("//table/tbody/tr/td[3]").Text;

                        Data data = new Data(cep, name, district, state);
                        Program.Write(data);
                        Console.WriteLine("Found CEP : " + cep);
                    }

                    browser.Navigate().Refresh();

                    outLoop = CEPcount + cepStart >= cepEnd; //Exit
                    if (outLoop)
                    {
                        Program.threads[id].Interrupt();
                        busy = false;
                        
                    }
                }
                catch(StaleElementReferenceException e)
                {
                    return;
                }
                catch (Exception ex) 
                {
                    if (ex is NoSuchElementException || ex is ElementNotInteractableException || ex is WebDriverTimeoutException)
                    {                        
                        browser.Navigate().Refresh();                            
                    }                   
                    else
                    {
                        busy = false;
                        throw ex;
                    }
                }
                            
            }
            
        }
    }
    public enum Fields
    {
        CEP = 1,
        Logradouro_Nome,
        Bairro_Distrito,
        Localidade_UF,
        Time_Added
    }
    public class Data // class to hold information
    {        
        public string cep;
        public string nome;
        public string bairro;
        public string estado;
        public DateTime timestamp;
        
        public Data(string _cep, string _nome, string _bairro, string _estado)
        {            
            cep = _cep;
            nome = _nome;
            bairro = _bairro;
            estado = _estado;
            timestamp = DateTime.Now;
        }
    }
}
