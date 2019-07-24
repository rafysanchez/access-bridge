using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.Events;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WindowsFormsTeste.PlugIn
{
    class TRJPlugin
    {
        IWebDriver _driver;
        EventFiringWebDriver _driverEvent;

        public delegate void MessageHandler(string message);
        public event MessageHandler MessageEventHandler;

        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public TRJPlugin BuildProcess()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--disable-notifications");

            string path = System.Environment.CurrentDirectory;

            _driver = new ChromeDriver(path, options);

            _driver.Manage().Timeouts().ImplicitWait = new System.TimeSpan(30);
            _driver.Manage().Window.Maximize();

            _driverEvent = new EventFiringWebDriver(_driver);

            return this;
        }

        public TRJPlugin OpenWebProcess(string url)
        {
            _log.Info("Passo OpenBizTalk Process....");
         
            _driverEvent.Navigate().GoToUrl(url);

            //SetDialogRPA();

            _log.Info("Passo OpenBizTalk ....OK");

            return this;
        }

        public string RetornaForo_CEP(string cep)
        {
            string foro = string.Empty;

            SelectElement tipoParcela = new SelectElement(_driver.FindElement(By.Id("optionsPesquisa")));

            foreach (IWebElement element in tipoParcela.Options)
            {
                if (element.Text.Contains("Cep"))
                {
                    element.Click();
                }
            }

            SearchAndReturnByID("inputPesquisa",20).SendKeys(cep);
            SearchAndReturnByID("BuscarButton", 10).Click();

            try
            {
                SearchAndReturnByClassName("k-grid-content", 10);

                var divTable = _driver.FindElement(By.ClassName("k-grid-content"));

                var ta = divTable.FindElement(By.TagName("table"));

                var taRow = ta.FindElements(By.TagName("tr"));

                foreach (var td in taRow)
                {
                    var cel = td.FindElements(By.TagName("td"));

                    foro = cel[3].Text;

                    break;
                }
            }
            catch { }

            return foro;
        }

        public void TakeScreenshot()
        {
            try
            {
                string path = @"%userprofile%\Downloads\";
                var filePath = Environment.ExpandEnvironmentVariables(path);


                string fileName = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".jpg";

                Byte[] byteArray = ((ITakesScreenshot)_driver).GetScreenshot().AsByteArray;
                System.Drawing.Bitmap screenshot = new System.Drawing.Bitmap(new System.IO.MemoryStream(byteArray));
                System.Drawing.Rectangle croppedImage = new System.Drawing.Rectangle(600, 100, 600, 140);
                screenshot = screenshot.Clone(croppedImage, screenshot.PixelFormat);
                screenshot.Save(String.Format(filePath + fileName, System.Drawing.Imaging.ImageFormat.Jpeg));


            }
            catch (Exception e)
            {
                //logger.Error(e.StackTrace + ' ' + e.Message);
            }

        }

        public TRJPlugin BaixarDocumento_FluxoProtesto(string doc)
        {
            bool documentoOk = false;
            bool documentoBase = false;

            try
            {
                var elTable = _driver.FindElement(By.TagName("table"));

                var count = elTable.FindElements(By.TagName("tr")).Count();

                int c = Convert.ToInt32(count);


                if (c > 0)
                {
                    foreach(var cel in elTable.FindElements(By.TagName("tr")))
                    {
                        var celula = cel.FindElements(By.TagName("td"));
                        if (celula.Count > 0)
                        {
                            if (celula[1].Text == doc)
                            {
                                string motivo = celula[7].Text;
                                string status = celula[8].Text;

                                if (status.Length > 0 && status.ToUpper() == "NEG")
                                {
                                    if (motivo.Length > 0 && motivo.ToUpper() != "AUSENTE" && motivo.ToUpper() != "NÃO PROCURADO")
                                    {
                                        // abrir notificação realizar download 
                                        var lnk = celula[9].FindElement(By.TagName("a"));
                                        lnk.Click();

                                        Thread.Sleep(2000);

                                        TakeScreenshot();

                                        // Validar Endereõ com Copia Contrato- Se OK => Salvar 
                                        var lnkSalvar = celula[10].FindElement(By.TagName("a"));
                                        lnkSalvar.Click();
                                        documentoOk = true;

                                        // verificar se já esta na base 
                                        try
                                        {  // Check the presence of alert
                                            Thread.Sleep(2000);
                                            _driver.SwitchTo().Alert().Accept();
                                            Thread.Sleep(1000);

                                            documentoBase = true;
                                        }
                                        catch (NoAlertPresentException ex)
                                        {
                                            documentoBase = false;
                                        }

                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { }


            return this;
        }

        public TRJPlugin BaixarDocumento(string doc)
        {
            bool documentoOk = false;
            bool documentoBase = false;

                try {

                var elTable = _driver.FindElement(By.TagName("table"));

                var count = elTable.FindElements(By.TagName("tr")).Count();

                int c = Convert.ToInt32(count);

                var eltr = elTable.FindElements(By.TagName("tr")).ToList()[c-1];
                var elCelula = eltr.FindElements(By.TagName("td"));
                if (elCelula.Count > 5)
                {
                    //Validar documento 
                    if (elCelula[1].Text == doc)
                    {
                        documentoOk = true;

                        var lnk = elCelula[10].FindElement(By.TagName("a"));
                        lnk.Click();
                    }
                }
                    
                    if (documentoOk)
                    {
                        // verificar se já esta na base 
                        try
                        {  // Check the presence of alert
                            Thread.Sleep(2000);
                            _driver.SwitchTo().Alert().Accept();
                            Thread.Sleep(1000);

                            documentoBase = true;
                        }
                        catch (NoAlertPresentException ex)
                        {
                            documentoBase = false; 
                        }

                        return this;
                    }

                } catch { }
            

            return this;
        }



        public IWebElement SearchAndReturnByClassName(string Name, int seconds)
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(seconds))
                .Until(ExpectedConditions.ElementExists((By.ClassName(Name))));

            return wait;
        }

        public IWebElement SearchAndReturnByID(string ID, int seconds)
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(seconds))
                .Until(ExpectedConditions.ElementExists((By.Id(ID))));

            return wait;
        }

        public void Close()
        {
            _driver.Close();
            _driver.Quit();
        }
    }
}
