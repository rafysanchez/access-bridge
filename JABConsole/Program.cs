using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WindowsAccessBridgeInterop;
using System.Windows.Forms;


namespace JABConsole
{
    class Program
    {
        static void Main(string[] args)
        {

            Application.Run();

            //var WindowForm = new WindowsFormsTeste.Form1();

            

            //var testete = WindowForm.getQtdProcess();


            var jab = new JABController();

            
                jab = new JABController();
                jab.InicializarDriver();


            jab.CheckJVM();

            
            Task t = new Task(() =>
            {
                Thread.Sleep(4000);
                while (true)
                {
                    jab.CheckJVM();
                    Thread.Sleep(4000);
                }
            });

            t.Start();

            Console.ReadKey();

        }
    
    }
}
