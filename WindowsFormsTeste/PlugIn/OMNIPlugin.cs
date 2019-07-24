using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TestStack.White;
using TestStack.White.InputDevices;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.WindowStripControls;


namespace WindowsFormsTeste.PlugIn
{
    class OMNIPlugin
    {
        IWebDriver _driver;
        EventFiringWebDriver _driverEvent;

        public delegate void MessageHandler(string message);
        public event MessageHandler MessageEventHandler;

        private Application _application;
        public Application Application { get { return _application; } set { _application = value; } }

        private Window _window;
        public Window Window { get { return _window; } set { _window = value; } }


        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public OMNIPlugin BuildProcess()
        {
            
            string path = System.Environment.CurrentDirectory;

            _driver = new InternetExplorerDriver(path + @"\Drivers");


            _driver.Manage().Timeouts().ImplicitWait = new System.TimeSpan(30);
            _driver.Manage().Window.Maximize();

            _driverEvent = new EventFiringWebDriver(_driver);

            return this;
        }

        public OMNIPlugin OpenWebProcess(string url)
        {
            _log.Info("Passo OpenOmni Process....");
         
            _driverEvent.Navigate().GoToUrl(url);

            //SetDialogRPA();

            _log.Info("Passo OpenOmni ....OK");

            return this;
        }

        public OMNIPlugin RealizarLogin()
        {
            Thread.Sleep(2000);
            try
            {
                Thread.Sleep(4000);
                Keyboard.Instance.Enter("robo_sow");
                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.TAB);
                Keyboard.Instance.Enter("Oracle2019");
                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                string erro = ex.Message;
                
            }

            return this;
        }

        public void Close()
        {
            _driver.Close();
        }
    }
}
