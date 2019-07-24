using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsTeste.PlugIn
{
    public class Servicos
    {
        public DesktopService DesktopService;
     

        public Servicos()
        {
            WebService = new WebService(
                new WebDriverContextInfo()
                {
                    Path = System.String.Concat(System.Environment.CurrentDirectory, @"\Drivers"),
                    MaximizeWindow = false,
                    Browser = BrowserEnum.InternetExplorer,
                    Driver = DriverEnum.Selenium,
                    Timeout = 60,
                    Attempts = 10,
                    MaxAttempts = 10
                }
            );

            DesktopService = new DesktopService(
                new DesktopDriverContextInfo()
                {
                    Driver = DriverEnum.TestStack,
                    Timeout = 30,
                    Attempts = 10,
                    MaxAttempts = 10
                    //,Path= System.String.Concat(System.Environment.CurrentDirectory, @"\Drivers")
                }
            );

            OCRService = new OCRService(
                new OCRDriverContextInfo()
                {
                    Driver = DriverEnum.Sikuli,
                    Matching = 0.9,
                    Timeout = 60,
                    Attempts = 10,
                    MaxAttempts = 10
                }
            );
        }
    }
}
