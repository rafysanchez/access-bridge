using CodeGen.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsAccessBridgeInterop;
using TestStack.White;
using TestStack.White.InputDevices;
using TestStack.White.UIItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.WindowStripControls;
using System.Threading;
using System.Diagnostics;
using TestStack.White.UIItems.ListBoxItems;
using TestStack.White.UIItems.WPFUIItems;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Automation;

namespace WindowsFormsTeste.PlugIn
{

    public class JAVAAPI
    {
        AccessBridge _accessBridge = new AccessBridge();
        private Application _application;
        public Application Application { get { return _application; } set { _application = value; } }
        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        BizTalkPlugin biz = new BizTalkPlugin();
        TRJPlugin trj = new TRJPlugin();

        private Window _window;
        public Window Window { get { return _window; } set { _window = value; } }

        public void InicializarDriver()
        {
            _accessBridge.Initialize();

            _accessBridge.Initilized += _accessBridge_Initilized;
        }

        private void _accessBridge_Initilized(object sender, EventArgs e){}
    }

    public class CobrancaPlugIn
    {
        AccessBridge _accessBridge = new AccessBridge();
        private Application _application;
        public Application Application { get { return _application; } set { _application = value; } }
        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        BizTalkPlugin biz = new BizTalkPlugin();
        TRJPlugin trj = new TRJPlugin();

        private Window _window;
        public Window Window { get { return _window; } set { _window = value; } }

        public void JanelaDownloadIE()
        {
            Process[] p = Process.GetProcesses();
            Process proc = new Process();

            bool findProcess = true;
            int tEspera = 10;
            int contEspera = 0;
            string processName = string.Empty;

            while (findProcess)
            {
                if (contEspera >= tEspera)
                {
                    findProcess = false;
                    return;
                }

                contEspera++;

                for (int i = 0; i < p.Length; i++)
                {
                    if (p[i].MainWindowTitle.Contains("View Downloads - Internet Explorer"))
                    {
                        proc = p[i];
                        findProcess = false;
                        processName = p[i].MainWindowTitle;
                        break;
                    }
                }

                Thread.Sleep(1000);
            }

            var wApp = Application.Attach(proc);
            _window = wApp.GetWindow(processName);
        }

        public void JanelaPDFPrint()
        {
            Process[] p = Process.GetProcesses();
            Process proc = new Process();

            bool findProcess = true;
            int tEspera = 10;
            int contEspera = 0;
            string processName = string.Empty;

            while (findProcess)
            {
                if (contEspera >= tEspera)
                {
                    findProcess = false;
                    return;
                }

                contEspera++;

                for (int i = 0; i < p.Length; i++)
                {
                    if (p[i].MainWindowTitle.Contains("Adobe Reader"))
                    {
                        proc = p[i];
                        findProcess = false;
                        processName = p[i].MainWindowTitle;
                        break;
                    }
                }

                Thread.Sleep(1000);
            }

            var wApp = Application.Attach(proc);
            _window = wApp.GetWindow(processName);
        }

        public void JanelaDownloadExplorer()
        {
            //Downloads
            Process[] p = Process.GetProcesses();
            Process proc = new Process();

            bool findProcess = true;
            int tEspera = 10;
            int contEspera = 0;
            string processName = string.Empty;

            while (findProcess)
            {
                if (contEspera >= tEspera)
                {
                    findProcess = false;
                    return;
                }

                contEspera++;

                for (int i = 0; i < p.Length; i++)
                {
                    if(p[i].ProcessName == "explorer")
                    {
                        try
                        {
                            p[i].Kill();
                        }
                        catch { }
                     }

                    //if (p[i].MainWindowTitle.Contains("Downloads"))
                    //{
                    //    proc = p[i];
                    //    findProcess = false;
                    //    processName = p[i].MainWindowTitle;
                    //    break;
                    //}
                }

                Thread.Sleep(1000);
            }

           // var wApp = Application.Attach(proc);
           // _window = wApp.GetWindow(processName);
        }

        public void FecharJanela(string janela)
        {

            JanelaDownloadExplorer();
            //var b = _window.Get<Button>(SearchCriteria.ByAutomationId("Close"));
            //b.Click();

        }

        public string DownloadDocumento(string numeroContrato)
        {
            string nomeDocumento = string.Empty;
            bool documentoDownload = false;
            
            bool continuar = true;
            ListItem Itens = null;
            int qtCorrente = 0;
            int qtMax = 5;

            while (continuar)
            {
                try
                {
                    if (qtCorrente <= qtMax)
                    {
                        Thread.Sleep(1000);
                        // 1 - Clicar em Abrir 
                        Itens = _window.Get<ListItem>();
                        continuar = false;
                    }

                }
                catch { qtCorrente++; }
            }

            if (Itens != null && Itens.Text.Contains(numeroContrato))
            {
                nomeDocumento = Itens.Text;
                documentoDownload = false;

                var link = Itens.GetMultiple(SearchCriteria.All);
                // Primeiro Verificar se ja foi realizado Download 
                foreach (var i in link)
                {
                    if (i.Name == "Downloads")
                    {
                        documentoDownload = true;
                        i.Click();
                        // _window.Mouse.Location = new System.Windows.Point(Itens.Location.X + 500, Itens.Location.Y + 25);
                        // _window.Mouse.Click();
                    }
                }

                if (documentoDownload == false)
                {
                   var sp = Itens.GetElement(SearchCriteria.ByControlType(System.Windows.Automation.ControlType.SplitButton));
                   var name = sp.Current.Name;
                   
                    var testeall = sp.FindAll(TreeScope.Subtree, Condition.TrueCondition);
                    var dteste = testeall[0].Current.Name;
                    var dteste1 = testeall[1];

                    var button = sp.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "split button"));

                    var pexists = sp.GetSupportedPatterns();


                    ((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
                    Thread.Sleep(500);
                    _window.Keyboard.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                    bool aguardarScan = true;

                    while (aguardarScan)
                    {
                        link = Itens.GetMultiple(SearchCriteria.All);

                        foreach (var i in link)
                        {
                            if (i.Name == "Downloads")
                            {
                                try
                                {
                                    i.Click();

                                    aguardarScan = false;
                                    documentoDownload = true;
                                }
                                catch { }
                                

                                break;
                                // _window.Mouse.Location = new System.Windows.Point(Itens.Location.X + 520, Itens.Location.Y + 25);
                                //_window.Mouse.Click();
                            }
                        }
                    }
                }
            }

            if (!documentoDownload)
                nomeDocumento = string.Empty;

            _window.Close();

            return nomeDocumento;
        }


        public void InicializarDriver()
        {
            _accessBridge.Initialize();

            _accessBridge.Initilized += _accessBridge_Initilized;
        }

        private void _accessBridge_Initilized(object sender, EventArgs e)
        {

        }

        public AccessibleJvm BuscarJanelaJava(string TituloJanela) {

            var qtdProcesso = _accessBridge.EnumJvms();

            AccessibleJvm jvm = null;

            foreach (var p in qtdProcesso)
            {
                var title = p.GetTitle();

                if (title.Contains("Omni"))
                {
                    jvm = p;

                    break;
                }
            }

            return jvm;

        }

        #region Fluxo Abertura .....

        public void AcessarMenuContratoParaMontagem()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var title1 = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[1].GetTitle();

                        var title = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetTitle();

                        var objeto = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[1];

                        //var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Contratos Para Montagem"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                var ret = _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                break;
                            }
                        }
                    }
                }
            }
        }

        public string PegarNomeComarca()
        {
            string comarca = string.Empty;

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {

                        var noInterface = BuscaInterface("text: Comarca:", n);

                        var cnode = (AccessibleContextNode)noInterface;
                        var prop = cnode.GetProperties(PropertyOptions.AccessibleText);


                        var acessible = (PropertyGroup)prop[0];
                        var content = (PropertyGroup)acessible.Children[8];
                        var c = content.Children[0].Value;

                        if (c.ToString() == "No text contents")
                        {
                            comarca = string.Empty;
                        }
                        else
                        {
                            comarca = c.ToString();
                        }

                        return comarca;
                    }
                }
            }

            return comarca;
        }


        public bool VerificarContratoOKParaAbertura()
        {
            bool contratoOkAbertura = false;
            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {

                                var noInterface = BuscaInterface("Motivo:", n);

                                var cnode = (AccessibleContextNode)noInterface;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);


                                var acessible = (PropertyGroup)prop[0];
                                var content = (PropertyGroup)acessible.Children[8];
                                var c = content.Children[0].Value;

                                if (c.ToString() == "No text contents")
                                {
                                    contratoOkAbertura = true;
                                }
                                else
                                {
                                    contratoOkAbertura = false;
                                }

                                return contratoOkAbertura;
                            }
                }
            }

            return contratoOkAbertura;
        }

        public string PegarNomeCliente(int index)
        {
            string Nome = string.Empty;
            int indexNome = index + 32;

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("page tab: Montagem / Remontagem Agente", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        foreach (var ccnnoodde in noInterface.GetChildren())
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Cliente"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexNome) {
                                    var acessible = (PropertyGroup)prop[0];
                                    var content = (PropertyGroup)acessible.Children[8];
                                    var c = content.Children[0].Value;

                                    return Nome = c.ToString();
                                }
                                else{}
                            }
                        }
                    }
                }
            }

            return Nome;

        }


        public string D_PegarNomeCliente(int index)
        {
            string Nome = string.Empty;
            int indexNome = index + 34;

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("page tab: Montagem / Remontagem Agente", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        foreach (var ccnnoodde in noInterface.GetChildren())
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Cliente"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexNome)
                                {
                                    var acessible = (PropertyGroup)prop[0];
                                    var content = (PropertyGroup)acessible.Children[8];
                                    var c = content.Children[0].Value;

                                    return Nome = c.ToString();
                                }
                                else { }
                            }
                        }
                    }
                }
            }

            return Nome;

        }

        public void D_GerarPastaRede(int index)
        {
            string Nome = string.Empty;
            int indexNome = index + 255;

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("page tab: Montagem / Remontagem Agente", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        foreach (var ccnnoodde in noInterface.GetChildren())
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Grava Documentos para Distribuição"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexNome)
                                {

                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                                    int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                    return;
                                }
                                else { }
                            }
                        }
                    }
                }
            }
        }


        public string ClicarAbertura(int index)
        {
            string Nome = string.Empty;
            int indexNome = index + 129;

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0]
                            .GetChildren();//.ToList()[3].GetChildren().ToList()[0].GetChildren();

                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Busca Documentos para"))
                            {
                                var teste = "";



                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexNome)
                                {

                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                                    int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                                }
                                else { }
                            }
                        }
                    }
                }
            }

            return Nome;
        }

        public string EnvioMontagem(int index)
        {
            string Nome = string.Empty;
            int indexNome = index + 145;

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0]
                            .GetChildren();//.ToList()[3].GetChildren().ToList()[0].GetChildren();

                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("panel"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexNome)
                                {

                                    var check = cnode.GetChildren().ToList()[0];
                                    var checkContexto = (AccessibleContextNode)check;


                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, checkContexto.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                                    int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, checkContexto.AccessibleContextHandle, ref ad1, out falha1);

                                    Thread.Sleep(1000);

                                    Desktop.Instance.Keyboard.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);


                                }
                                else { }
                            }
                        }
                    }
                }
            }

            return Nome;
        }

        public void LerTabelaBuscaApreensao(string tipo, string digital)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach(var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0]
                            .GetChildren();//.ToList()[3].GetChildren().ToList()[0].GetChildren();

                        List<string> contratos = new List<string>();


                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("text: Contrato"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int index = cnode.GetIndexInParent();
                                var acessible = (PropertyGroup)prop[0];
                                var content = (PropertyGroup)acessible.Children[8];
                                var c = content.Children[0].Value;

                                if (c.ToString() == "No text contents")
                                {
                                    continue;
                                }
                                else
                                {
                                    var contentPoint = (PropertyGroup)acessible.Children[7];
                                    var point = contentPoint.Children[0];
                                    var valuePoint = (AccessibleRectInfo)point.Value;
                                    Desktop.Instance.Mouse.Click(new System.Windows.Point(valuePoint.X, valuePoint.Y));

                                    if (VerificarContratoOKParaAbertura())
                                    {

                                        contratos.Add(c.ToString());
                                        // Pegar Nome 
                                        string nome = PegarNomeCliente(index);
                                        bool birigui = false;

                                        string Comarca = PegarNomeComarca();

                                        if (Comarca.ToUpper().Contains("BIRIGUI"))
                                        {
                                            birigui = true;
                                        }


                                        // Clicar Abrir Tela do Contrato.
                                        ClicarAbertura(index);


                                        if (tipo == "N")
                                        {
                                            if (digital == "S")
                                            {
                                                Fluxo_Abertura_Notificacao_Digital(c.ToString(), nome);
                                            }
                                            else
                                            {
                                                Fluxo_Abertura_Notificacao_Fisica(c.ToString(), nome);
                                            }
                                        }
                                        else
                                        {
                                            if (digital == "S")
                                            {
                                                Fluxo_Abertura_Protesto_Digital(c.ToString(), nome, birigui);
                                            }
                                            else
                                            {
                                                Fluxo_Abertura_Protesto_Fisica(c.ToString(), nome, birigui);
                                            }
                                        }

                                        Thread.Sleep(2000);

                                        this.AbrirDocumento_DistribuicaoProcesso(index);

                                        this.EnvioMontagem(index);

                                        this.SelecionarBotaoAbrirProcesso();

                                        var demo = "Homologação";

                                        //  return;

                                    }

                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                                    int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                }
                                //break;
                            }
                        }
                    }
                }
            }
        }

        public void ScroolTabelaBuscaApreensao()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0]
                            .GetChildren();//.ToList()[3].GetChildren().ToList()[0].GetChildren();

                        List<string> contratos = new List<string>();


                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (tititi.Contains("scroll bar"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);

                                
                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                       

                                var acessible = (PropertyNode)prop[7];
                                var contentPoint = (AccessibleRectInfo)acessible.Value;

                                for(int i= 0; i<16; i++) {

                                    Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X+2, contentPoint.Y + contentPoint.Height-5));

                                }

                                int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                    return;
                                                           }

                        }
                    }
                }
            }
        }

        public void ClicardialogBoxAviso()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        bool novamente = true;
                        AccessibleNode noInterface = null;

                        while (novamente)
                        {
                            noInterface = BuscaInterface("button: OK", n);

                            if (noInterface != null)
                                novamente = false;
                        }

                        var cnode = (AccessibleContextNode)noInterface;

                        AccessibleActions acc = null;
                        _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                        AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                        {
                            actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                            actionsCount = 1
                        };

                        ad1.actions[0].name = "Click";

                        int falha1 = 0;

                        _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                    }
                }
            }

            Thread.Sleep(1000);
        }

        public void SelecionarBotaoAbrirProcesso()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0]
                            .GetChildren();//.ToList()[3].GetChildren().ToList()[0].GetChildren();


                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("button: AbrirProcesso"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);


                                break;
                            }
                        }
                    }
                }

            }
        }

        public void SelecionarBotaoPesquisar()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0]
                            .GetChildren();//.ToList()[3].GetChildren().ToList()[0].GetChildren();


                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("button: Pesquisar"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);


                                break;
                            }
                        }
                    }
                }

            }
        }

        public void SelecionarDigital(string tipo)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0]
                            .GetChildren();//.ToList()[3].GetChildren().ToList()[0].GetChildren();


                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Processo Digital"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;

                                var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                                var acessible = (PropertyNode)prop[7];

                                var contentPoint = (AccessibleRectInfo)acessible.Value;
                                Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X, contentPoint.Y));
                                Keyboard.Instance.Enter(tipo);
                                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                                break;
                            }
                        }
                    }
                }
            }
        }

        public void SelecionarAdvogado(string nome)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0]
                            .GetChildren();//.ToList()[3].GetChildren().ToList()[0].GetChildren();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Advogado"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;

                                var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                                var acessible = (PropertyNode)prop[7];

                                var contentPoint = (AccessibleRectInfo)acessible.Value;
                                Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X, contentPoint.Y));

                                if (nome == "0")
                                {

                                    Keyboard.Instance.Enter("9");
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.UP);
                                }
                                else
                                {
                                    Keyboard.Instance.Enter("9");
                                }

                                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                                break;
                            }
                        }
                    }
                }

            }
        }

        public void SelecionarTipoMontagem(string tipo)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0]
                            .GetChildren();//.ToList()[3].GetChildren().ToList()[0].GetChildren();

                        var title = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetTitle();


                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Tipo Montagem"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                                var acessible = (PropertyNode)prop[7];


                                var contentPoint = (AccessibleRectInfo)acessible.Value;
                                Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X, contentPoint.Y));
                                Keyboard.Instance.Enter(tipo);
                                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                                break; ;
                            }
                        }
                    }
                }
            }
        }

        public void BuscarProcessos(string nomeProcesso)
        {
            var qtdProcesso = _accessBridge.EnumJvms();

            AccessibleJvm jvm = null;

            foreach (var p in qtdProcesso)
            {
                var title = p.GetTitle();

                if (title.Contains("Omni"))
                {
                    jvm = p;

                    break;
                }
            }


            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var title1 = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[1].GetTitle();

                        var title = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetTitle();

                        var objeto = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[1];

                        //var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren();



                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Contratos Para Montagem"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                var ret = _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                break;
                            }
                        }
                    }

                }

            }

        }

        public void ClicarRetornarTelaPrincipal()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var veri = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();


                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("button: Retorna a tela principal"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                Thread.Sleep(2000);

                            }
                        }
                    }
                }
            }
        }

        public void AbrirDocumento_DistribuicaoProcesso(int indexLupa) {

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var veri = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();


                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Exibe o documento digitalizado"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexLupa)
                                {
                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                                    int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                                }
                                else
                                {

                                }

                            }
                        }
                    }
                }
            }
        }

        public void ValidarCheckDocumento_DistribuicaoProcesso(int indexCheck)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var veri = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();


                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Efetua a validação do documento"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexCheck)
                                {
                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                                    int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                    Thread.Sleep(3000);
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.TAB);
                                    Thread.Sleep(1000);
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);


                                    try
                                    {
                                        Thread.Sleep(1500);
                                        D_BotaoDadosGravadosSucesso(5);
                                    }
                                    catch { }


                                }
                            }
                        }
                    }
                }
            }
        }

        public void D_BotaoDadosGravadosSucesso(int tentativas)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        List<string> contratos = new List<string>();

                        var noInterface = BuscaInterface("Aviso Dados gravados com sucesso!", n).GetChildren().ToList()[0].GetChildren().ToList()[0];

                        foreach (var ccnnoodde in noInterface.GetChildren())
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("OK"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);


                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                return;


                            }
                        }
                    }
                }
            }
        }

        public void ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(int indexCheck)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var veri = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();


                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("text: Seq."))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexCheck)
                                {
                                    var acessible = (PropertyGroup)prop[0];
                                    var contentPoint = (PropertyGroup)acessible.Children[7];
                                    var point = contentPoint.Children[0];
                                    var valuePoint = (AccessibleRectInfo)point.Value;
                                    Desktop.Instance.Mouse.Click(new System.Windows.Point(valuePoint.X + 3, valuePoint.Y + 2));

                                    return;
                                }
                            }
                        }
                    }
                }
            }
        }

        public bool ValidarDocumentoProtestoSolicitadoCancelado_DistribuicaoProcesso(int indexDigital)
        {
            bool documentoX = false; ;

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var veri = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();


                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Status Solicitação"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexDigital)
                                {
                                    var acessible = (PropertyGroup)prop[0];
                                    var content = (PropertyGroup)acessible.Children[8];
                                    var c = content.Children[0].Value;

                                    if (c.ToString().ToUpper().Contains("CANCELAMENTO"))
                                        return documentoX = true;
                                }
                                else
                                {

                                }

                            }
                        }
                    }
                }
            }
            return documentoX;
        }

        public bool ValidarDocumentoProtestojaValidado_DistribuicaoProcesso(int indexDigital)
        {
            bool documentoX = false; ;

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var veri = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();


                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Dt. Validação"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexDigital)
                                {
                                    var acessible = (PropertyGroup)prop[0];
                                    var content = (PropertyGroup)acessible.Children[8];
                                    var c = content.Children[0].Value;

                                    if (!c.ToString().Contains("No text contents"))
                                        return documentoX = true;
                                }
                                else
                                {

                                }

                            }
                        }
                    }
                }
            }
            return documentoX;
        }

        public bool ValidarDocumentoDigital_DistribuicaoProcesso(int indexDigital)
        {
            bool documentoX = false; ;

            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var veri = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();


                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Digital"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexDigital)
                                {
                                    var acessible = (PropertyGroup)prop[0];
                                    var content = (PropertyGroup)acessible.Children[8];
                                    var c = content.Children[0].Value;

                                    if (c.ToString() == "X")
                                        return documentoX = true;
                                }
                                else
                                {

                                }

                            }
                        }
                    }
                }
            }
            return documentoX;
        }

        public void SolicitaDocumento_DistribuicaoProcesso(int indexDigital)
        {
            //Desconsiderar Contratos Linha Vermelha e Amarela 
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var veri = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();


                        List<string> contratos = new List<string>();

                        int co = ot.Count();

                        foreach (var ccnnoodde in ot)
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("Solicita documento"))
                            {
                                var teste = "";

                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int ind = cnode.GetIndexInParent();

                                if (ind == indexDigital)
                                {
                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                                    int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                    Thread.Sleep(3000);
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.TAB);
                                    Thread.Sleep(1000);
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);


                                    Thread.Sleep(2000);
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                                    // this.ClicardialogBoxAviso();

                                }
                                else
                                {

                                }

                            }
                        }
                    }
                }
            }
        }

        public void Fluxo_Abertura_Protesto_Digital(string contrato, string nome, bool birigui)
        {

            bool CopiaContrato = false;
            bool Notificacao = false;
            bool InstrumentoProtesto = false;
            bool ProtestoSolicitacaoCanelada = false;

            bool CopiaContratoValidado = false;
            bool ProtestoValidado = false;
            bool NotificacaoValidado = false;

            bool DocNotificacao = false;


            CopiaContratoValidado = ValidarDocumentoProtestojaValidado_DistribuicaoProcesso(144);
            ProtestoValidado = ValidarDocumentoProtestojaValidado_DistribuicaoProcesso(145);
            NotificacaoValidado = ValidarDocumentoProtestojaValidado_DistribuicaoProcesso(146);

            //1 - Verificar se Digital esta com X 
            CopiaContrato = ValidarDocumentoDigital_DistribuicaoProcesso(88);
            Notificacao = ValidarDocumentoDigital_DistribuicaoProcesso(90);
            InstrumentoProtesto = ValidarDocumentoDigital_DistribuicaoProcesso(89);
            ProtestoSolicitacaoCanelada = ValidarDocumentoProtestoSolicitadoCancelado_DistribuicaoProcesso(75);

            _log.InfoFormat("Abertura Digital : {0} {1}", nome, contrato);

            if (!CopiaContratoValidado)
            {
                if (!CopiaContrato)
                {
                    this.SolicitaDocumento_DistribuicaoProcesso(158);
                }
                else
                {
                    //Validar Documento - Abrir Gerar OCR Capturar Nome e Endereço 
                    this.AbrirDocumento_DistribuicaoProcesso(186);

                    Thread.Sleep(2000);

                    // Realizar Download e OCR 
                    DownloadArquivoIE(contrato);

                }
            }

            if (!ProtestoValidado)
            {
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(5);

                if (InstrumentoProtesto)
                {
                    //Abrir Instrumento ??
                    this.AbrirDocumento_DistribuicaoProcesso(187);

                    Thread.Sleep(2000);
                    // Realizar Download e OCR 
                    DownloadArquivoIE(contrato);
                }
                else
                {
                    //if (!ProtestoSolicitacaoCanelada)
                    //    this.SolicitaDocumento_DistribuicaoProcesso(159);

                    //Enviar Fila de Exceção
                }
            }

            if (!NotificacaoValidado)
            {
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(6);

                if (!Notificacao)
                {
                    biz.BuildProcess()
                       .OpenWebProcess((string.Format("http://www.bizfacil.com.br/pls/webdad/pck_consulta_cartorio.prc_consulta_cartorio?pContrato={0}", contrato)));

                    DocNotificacao = biz.BaixarDocumento_FluxoProtesto(contrato, birigui);

                    biz.Close();


                }
                else
                {
                    //Validar Documento
                }
            }

            if (!CopiaContratoValidado)
            {
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(4);
                Thread.Sleep(500);
                this.ValidarCheckDocumento_DistribuicaoProcesso(130);
            }
            Thread.Sleep(1000);

            if (!ProtestoValidado)
            {
                if (InstrumentoProtesto)
                {
                    ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(5);
                    this.ValidarCheckDocumento_DistribuicaoProcesso(131);
                }
            }

            if (!NotificacaoValidado)
            {
                //Notificacao = ValidarDocumentoDigital_DistribuicaoProcesso(90);

                //if (Notificacao)
                //{
                    Thread.Sleep(500);
                    ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(5);
                    ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(6);
                    this.ValidarCheckDocumento_DistribuicaoProcesso(132);
                //}


                Thread.Sleep(1000);
            }

            this.ClicarRetornarTelaPrincipal();
            var testeNot = "";


        }

        public void Fluxo_Abertura_Protesto_Fisica(string contrato, string nome, bool birigui)
        {
            bool CopiaContrato = false;
            bool Notificacao = false;
            bool InstrumentoProtesto = false;
            bool ProtestoSolicitacaoCanelada = false;

            bool CopiaContratoValidado = false;
            bool ProtestoValidado = false;
            bool NotificacaoValidado = false;

            bool DocNotificacao = false;


            CopiaContratoValidado = ValidarDocumentoProtestojaValidado_DistribuicaoProcesso(144);
            ProtestoValidado = ValidarDocumentoProtestojaValidado_DistribuicaoProcesso(145);
            NotificacaoValidado = ValidarDocumentoProtestojaValidado_DistribuicaoProcesso(146);

            //1 - Verificar se Digital esta com X 
            CopiaContrato = ValidarDocumentoDigital_DistribuicaoProcesso(88);
            Notificacao = ValidarDocumentoDigital_DistribuicaoProcesso(90);
            InstrumentoProtesto = ValidarDocumentoDigital_DistribuicaoProcesso(89);
            ProtestoSolicitacaoCanelada = ValidarDocumentoProtestoSolicitadoCancelado_DistribuicaoProcesso(75);

            _log.InfoFormat("Abertura Digital : {0} {1}", nome, contrato);

            if (!CopiaContratoValidado)
            {
                if (!CopiaContrato)
                {
                    this.SolicitaDocumento_DistribuicaoProcesso(158);
                }
                else
                {
                    //Validar Documento - Abrir Gerar OCR Capturar Nome e Endereço 
                    this.AbrirDocumento_DistribuicaoProcesso(186);

                    Thread.Sleep(2000);

                    // Realizar Download e OCR 
                    DownloadArquivoIE(contrato);

                }
            }

            if (!ProtestoValidado)
            {
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(5);

                if (InstrumentoProtesto)
                {
                    //Abrir Instrumento ??
                    this.AbrirDocumento_DistribuicaoProcesso(187);

                    Thread.Sleep(2000);
                    // Realizar Download e OCR 
                    DownloadArquivoIE(contrato);
                }
                else
                {
                    //if (!ProtestoSolicitacaoCanelada)
                    //    this.SolicitaDocumento_DistribuicaoProcesso(159);

                    //Enviar Fila de Exceção
                }
            }

            if (!NotificacaoValidado)
            {
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(6);

                if (!Notificacao)
                {
                    biz.BuildProcess()
                       .OpenWebProcess((string.Format("http://www.bizfacil.com.br/pls/webdad/pck_consulta_cartorio.prc_consulta_cartorio?pContrato={0}", contrato)));

                    DocNotificacao = biz.BaixarDocumento_FluxoProtesto(contrato, birigui);

                    biz.Close();


                }
                else
                {
                    //Validar Documento
                }
            }

            if (!CopiaContratoValidado)
            {
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(4);
                Thread.Sleep(500);
                this.ValidarCheckDocumento_DistribuicaoProcesso(130);
            }
            Thread.Sleep(1000);

            if (!ProtestoValidado)
            {
                if (InstrumentoProtesto)
                {
                    ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(5);
                    this.ValidarCheckDocumento_DistribuicaoProcesso(131);
                }
            }

            if (!NotificacaoValidado)
            {
                //Notificacao = ValidarDocumentoDigital_DistribuicaoProcesso(90);

                //if (Notificacao)
                //{
                Thread.Sleep(500);
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(5);
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(6);
                this.ValidarCheckDocumento_DistribuicaoProcesso(132);
                //}


                Thread.Sleep(1000);
            }


            //Imprimir Documentos 

            LerArquivoImprimir(contrato);


            Thread.Sleep(1000);

            this.ClicarRetornarTelaPrincipal();
            var testeNot = "";

        }

        public void Fluxo_Abertura_Notificacao_Digital(string contrato, string nome)
        {
            bool CopiaContrato = false;
            bool Notificacao = false;
            //1 - Verificar se Digital esta com X 
            CopiaContrato = ValidarDocumentoDigital_DistribuicaoProcesso(88);
            Notificacao = ValidarDocumentoDigital_DistribuicaoProcesso(89);

            _log.InfoFormat("Abertura Digital : {0} {1}", nome, contrato);

            if (!CopiaContrato) {

                this.SolicitaDocumento_DistribuicaoProcesso(158);
            }


            if (!Notificacao)
            {
                biz.BuildProcess()
                    .OpenWebProcess((string.Format("http://www.bizfacil.com.br/pls/webdad/pck_consulta_cartorio.prc_consulta_cartorio?pContrato={0}", contrato)))
                    .BaixarDocumento(contrato)
                    .Close();

            }

                this.ValidarCheckDocumento_DistribuicaoProcesso(130);

                Thread.Sleep(2000);

                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(5);
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(4);
                ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(5);

                this.ValidarCheckDocumento_DistribuicaoProcesso(131);


            this.ClicarRetornarTelaPrincipal();
            var testeNot = "";


        }

        public void Fluxo_Abertura_Notificacao_Fisica(string contrato, string nome)
        {
            bool CopiaContrato = false;
            bool Notificacao = false;
            //1 - Verificar se Digital esta com X 
            CopiaContrato = ValidarDocumentoDigital_DistribuicaoProcesso(88);
            Notificacao = ValidarDocumentoDigital_DistribuicaoProcesso(90);
            bool DocNotificacao = false;
           
            string pathArquivoContrato = string.Empty;

            _log.InfoFormat("Abertura Digital : {0} {1}", nome, contrato);

            if (!CopiaContrato)
            {
                this.SolicitaDocumento_DistribuicaoProcesso(158);
            }
            else
            {
                //Validar Documento - Abrir Gerar OCR Capturar Nome e Endereço 
                this.AbrirDocumento_DistribuicaoProcesso(186);

                Thread.Sleep(2000);
                // Realizar Download e OCR 
                pathArquivoContrato = DownloadArquivoIE(contrato);

            }


            if (!Notificacao)
            {
                biz.BuildProcess()
                    .OpenWebProcess((string.Format("http://www.bizfacil.com.br/pls/webdad/pck_consulta_cartorio.prc_consulta_cartorio?pContrato={0}", contrato)));

                   DocNotificacao = biz.BaixarDocumento_FluxoProtesto(contrato, false);

                    biz.Close();
            }
            else
            {
                //Validar Documento
            }


            this.ValidarCheckDocumento_DistribuicaoProcesso(130);

            Thread.Sleep(2000);

            //D_BotaoDadosGravadosSucesso(2);
            ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(5);
            if (!Notificacao)
            {
                if(DocNotificacao)
                    this.ValidarCheckDocumento_DistribuicaoProcesso(131);
            }
            else
            {
                this.ValidarCheckDocumento_DistribuicaoProcesso(131);
            }

            Thread.Sleep(2000);
            //ColocaFocoLinhaValidacaoDocumento_DistribuicaoProcesso(6);
            //this.ValidarCheckDocumento_DistribuicaoProcesso(132);

            //Imprimir Documentos 

            LerArquivoImprimir(pathArquivoContrato.Replace(".pdf", ""));




            Thread.Sleep(1000);

            this.ClicarRetornarTelaPrincipal();
            var testeNot = "";

        }

        #endregion

        private void ProcessoImprimir(string pathCompleto)
        {
            ProcessStartInfo info = new ProcessStartInfo();
            info.Verb = "print";
            info.FileName = pathCompleto; // "nome do arquivo.pdf";
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Hidden;
            info.UseShellExecute = true;

            Process p = new Process();
            p.StartInfo = info;
            p.Start();

            p.WaitForInputIdle();
            System.Threading.Thread.Sleep(3000);
            if (false == p.CloseMainWindow())
                p.Kill();
        }

        [DllImport("shell32.dll", EntryPoint = "ShellExecute")]
        public static extern int ShellExecuteA(int hwnd, string lpOperation, string lpFile, string lpParameters, string lpDirectory, int nShowCmd);

        public void LerArquivoImprimir(string nomeArquivo)
        {
            string path = @"%userprofile%\Downloads\";
            var filePath = Environment.ExpandEnvironmentVariables(path);

            Directory.EnumerateFiles(filePath).ToList().ForEach(delegate (string strfile)
            {
                if (strfile.Contains(nomeArquivo))
                {
                    if (File.Exists(strfile))
                    {
                        try
                        {
                            ProcessoImprimir(strfile);

                            File.Delete(strfile);

                        }
                        catch { }
                    }
                }

            });
        }


        private void LerArquivoGerarOcr(string nomeArquivo)
        {
            string path = @"%userprofile%\Downloads\";
            var filePath = Environment.ExpandEnvironmentVariables(path);

            if (File.Exists(filePath + nomeArquivo))
            {
                //Realizar OCR 

                File.Delete(filePath + nomeArquivo);
            }
        }


        private string DownloadArquivoIE(string contrato)
        {
            JanelaDownloadIE();

            string nomearquivo = DownloadDocumento(contrato);

            if (nomearquivo.Length > 0)
            {
                LerArquivoGerarOcr(nomearquivo);
            }


            FecharJanela("");
            //var b = _window.Get<Button>(SearchCriteria.ByAutomationId("Close"));
            //b.Click();

         //   Keyboard.Instance.HoldKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.CONTROL);
         //   Keyboard.Instance.Enter("W");
         //   Keyboard.Instance.LeaveKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.CONTROL);

            return nomearquivo;
        }



        #region Fluxo Distribuicao 

        public void D_Menu_MontageRemontagem()
        {

            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var ot = n.GetChildren().ToList()[1].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[8].GetChildren().ToList()[0].GetChildren().ToList()[2]
                            .GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();//.ToList()[2].GetChildren();

                        bool novamente = true;
                        AccessibleNode noInterface = null;

                        while (novamente)
                        {
                            noInterface = BuscaInterface("Montagem / Remontagem (DIGITAL)", n);

                            if (noInterface != null)
                                novamente = false;
                        }

                        var cnode = (AccessibleContextNode)noInterface;
                        var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                        var acessible = (PropertyNode)prop[7];

                        var contentPoint = (AccessibleRectInfo)acessible.Value;
                        Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X + 20, contentPoint.Y + 10));
                        
                    }
                }
            }
        }

        public void D_SelecionarAdvogado(string nome)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                            var noInterface = BuscaInterface("Advogado", n);

                                var cnode = (AccessibleContextNode)noInterface;

                                var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                                var acessible = (PropertyNode)prop[7];

                                var contentPoint = (AccessibleRectInfo)acessible.Value;
                                Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X, contentPoint.Y));

                                if (nome == "0")
                                {

                                    Keyboard.Instance.Enter("9");
                                    Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.UP);
                                }
                                else
                                {
                                    Keyboard.Instance.Enter("968");
                                }

                                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                                break;
                            
                        
                    }
                }

            }
        }

        public void D_SelecionarBotaoPesquisar()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                            var noInterface = BuscaInterface("button: Pesquisar", n);

                                var cnode = (AccessibleContextNode)noInterface;

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                    }
                }

            }
        }

        public void D_SelecionarTipoMontagem(string tipo)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("combo box: Tipo", n);

                        var cnode = (AccessibleContextNode)noInterface;
                        var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                        var acessible = (PropertyNode)prop[7];


                        var contentPoint = (AccessibleRectInfo)acessible.Value;
                        Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X, contentPoint.Y));
                        Keyboard.Instance.Enter(tipo);
                        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                    }
                }
            }
        }

        public void D_SelecionarTipomontagemNotificacao(string tipo)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("combo box: Tipo Montagem", n);

                        var cnode = (AccessibleContextNode)noInterface;
                        var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                        var acessible = (PropertyNode)prop[7];


                        var contentPoint = (AccessibleRectInfo)acessible.Value;
                        Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X, contentPoint.Y));
                        Keyboard.Instance.Enter(tipo);
                        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);
                    }
                }
            }
        }


        public void D_LerTabelaBuscaApreensao(string digital)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        List<string> contratos = new List<string>();

                        var noInterface = BuscaInterface("page tab: Montagem / Remontagem Agente", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        foreach (var ccnnoodde in noInterface.GetChildren())
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("text: Contrato"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int index = cnode.GetIndexInParent();
                                var acessible = (PropertyGroup)prop[0];
                                var content = (PropertyGroup)acessible.Children[8];
                                var c = content.Children[0].Value;

                                if (c.ToString() == "No text contents")
                                {
                                    continue;
                                }
                                else
                                {
                                    var contentPoint = (PropertyGroup)acessible.Children[7];
                                    var point = contentPoint.Children[0];
                                    var valuePoint = (AccessibleRectInfo)point.Value;
                                    Desktop.Instance.Mouse.Click(new System.Windows.Point(valuePoint.X, valuePoint.Y));

                                    if (!VerificarContratoOKParaAbertura())
                                    {
                                        contratos.Add(c.ToString());
                                        // Pegar Nome 
                                        string nome = D_PegarNomeCliente(index);

                                        Desktop.Instance.Mouse.DoubleClick(new System.Windows.Point(valuePoint.X, valuePoint.Y));

                                        var teste = "teste";

                                        // Clicar Abrir Tela do Contrato.
                                        Thread.Sleep(1500);
                                        D_LerOcorrenciasCadastroProcessoJuridico();

                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public void D_LerTabelaBuscaApreensao_CadastroForo(string digital)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        List<string> contratos = new List<string>();

                        var noInterface = BuscaInterface("page tab: Montagem / Remontagem Agente", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0];

                        foreach (var ccnnoodde in noInterface.GetChildren())
                        {
                            var tititi = ccnnoodde.GetTitle();

                            if (ccnnoodde.GetTitle().Contains("text: Contrato"))
                            {
                                var cnode = (AccessibleContextNode)ccnnoodde;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                int index = cnode.GetIndexInParent();
                                var acessible = (PropertyGroup)prop[0];
                                var content = (PropertyGroup)acessible.Children[8];
                                var c = content.Children[0].Value;

                                if (c.ToString() == "No text contents")
                                {
                                    continue;
                                }
                                else
                                {
                                    var contentPoint = (PropertyGroup)acessible.Children[7];
                                    var point = contentPoint.Children[0];
                                    var valuePoint = (AccessibleRectInfo)point.Value;
                                    Desktop.Instance.Mouse.Click(new System.Windows.Point(valuePoint.X, valuePoint.Y));

                                    if (VerificarContratoOKParaAbertura())
                                    {
                                        contratos.Add(c.ToString());
                                        // Pegar Nome 
                                        string nome = D_PegarNomeCliente(index);

                                        Desktop.Instance.Mouse.DoubleClick(new System.Windows.Point(valuePoint.X, valuePoint.Y));

                                        var teste = "teste";

                                        // Clicar Abrir Tela do Contrato.

                                        //Verificar Cidade 
                                        if (D_VerificaCidadeSP())
                                        {
                                            D_BotaoEndereco_Foro();
                                            string cep = D_PegaDadosCEP();

                                            Thread.Sleep(1000);

                                            string foro = D_BuscarDadosForoTRJ(cep);

                                            D_JanelaSelecionarForo();
                                            D_PreencherFiltroForo(foro);
                                            D_BotaoOKListaForo();
                                        }
                                        else
                                        {
                                            string uf = string.Empty;
                                            string cidade = string.Empty;
                                            string comarca = string.Empty;

                                            // Pegar Dados UF | Cidade | Comarca 
                                            D_PegaDadosUfCidadeComarca(out uf, out cidade,out comarca);
                                            //Se estiver Regra - Pesquisar Comarca e Inserir 

                                            //Se não estiver Regra - Selecionar Foro Central 
                                            // 1 - Cidade = Comarca -> Selecionar Foro Central 
                                            // 2 - Cidade |= Comarca -> Selecionar Foro Mesmo Nome da Cidade 
                                                if (cidade == comarca)
                                                {
                                                    D_JanelaSelecionarForo();
                                                    D_BotaoOKListaForo();
                                                    //D_BotaoSairProcessoJuridico();
                                                var sair = "";

                                            }
                                            else
                                                {
                                                    D_JanelaSelecionarForo();
                                                    D_PreencherFiltroForo(cidade);

                                                 if (D_RetornaRegistrosFiltroForo() == 0)
                                                 {
                                                    D_PreencherFiltroForo("-");
                                                 }

                                                D_BotaoOKListaForo();
                                                    //D_BotaoSairProcessoJuridico();
                                                var sair = "";
                                                }
                                        }

                                        //Gravar e Sair 
                                        D_BotaoSalvar();
                                        D_BotaoSairProcessoJuridico();

                                        Thread.Sleep(1000);
                                        D_GerarPastaRede(index);

                                        //Retornar Tela e Clicar Salvar 


                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public string D_BuscarDadosForoTRJ(string cep)
        {
            string foro = string.Empty;
           
            trj.BuildProcess()
            .OpenWebProcess("http://www.tjsp.jus.br/CompetenciaTerritorial");

            foro = trj.RetornaForo_CEP(cep);

            trj.Close();

            return foro;
        }

        public void D_PreencherFiltroForo(string filtro)
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("Lista de Foruns", n).GetChildren().ToList()[0].GetChildren().ToList()[1].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("text: "))
                            {
                                var cnode = (AccessibleContextNode)noInterface;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                var acessible = (PropertyGroup)prop[0];
                                var contentPoint = (PropertyGroup)acessible.Children[7];
                                var point = contentPoint.Children[0];
                                var valuePoint = (AccessibleRectInfo)point.Value;
                                Desktop.Instance.Mouse.Click(new System.Windows.Point(valuePoint.X, valuePoint.Y));
                                Desktop.Instance.Mouse.DoubleClick(new System.Windows.Point(valuePoint.X, valuePoint.Y));

                                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DELETE);
                                Keyboard.Instance.Enter("%" + filtro);

                                Thread.Sleep(1000);
                                D_BotaoProcurarForo();
                                Thread.Sleep(1000);

                                return;
                            }
                        }
                    }
                }

            }
        }

        public int D_RetornaRegistrosFiltroForo()
        {
            int qtd = 0;
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("Lista de Foruns", n).GetChildren().ToList()[0].GetChildren().ToList()[2].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("list"))
                            {
                                var cnode = (AccessibleContextNode)noInterface;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);

                                var acessible = (PropertyNode)prop[8];
                                var content = acessible.Value;

                                qtd = 0;
                                Int32.TryParse(content.ToString(), out qtd);

                                return qtd;
                            }
                        }
                    }
                }

            }

            return qtd;
        }


        public void D_JanelaSelecionarForo()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("page tab: Processos", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var cnode = (AccessibleContextNode)noInterface;
                            int ind = cnode.GetIndexInParent();

                            if (ind == 24)
                            {
                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                return;
                            }
                        }
                    }
                }
            }
        }

        public void D_PegaDadosUfCidadeComarca(out string uf, out string cidade, out string comarca)
        {
            uf = cidade = comarca = string.Empty;

            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("page tab: Processos", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var cnode = (AccessibleContextNode)noInterface;
                            int ind = cnode.GetIndexInParent();

                            if (ind == 14)
                            {
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);
                                var acessible = (PropertyGroup)prop[0];
                                var content = (PropertyGroup)acessible.Children[8];
                                var c = content.Children[0].Value;

                                uf = c.ToString().ToUpper();
                            }

                            if (ind == 16)
                            {
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);
                                var acessible = (PropertyGroup)prop[0];
                                var content = (PropertyGroup)acessible.Children[8];
                                var c = content.Children[0].Value;

                                cidade = c.ToString().ToUpper();
                            }

                            if (ind == 19)
                            {
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);
                                var acessible = (PropertyGroup)prop[0];
                                var content = (PropertyGroup)acessible.Children[8];
                                var c = content.Children[0].Value;

                                comarca = c.ToString().ToUpper();
                            }
                        }
                    }   
                }
            }

        }

        public string D_PegaDadosCEP()
        {
            string cep = string.Empty;
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("internal frame: Endereços", n).GetChildren().ToList()[0].GetChildren().ToList()[2].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {

                            var t = noInterface.GetTitle();
                            var cnode = (AccessibleContextNode)noInterface;
                            int ind = cnode.GetIndexInParent();

                            if (ind == 2)
                            {
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);
                                var acessible = (PropertyGroup)prop[0];
                                var content = (PropertyGroup)acessible.Children[8];
                                var c = content.Children[0].Value;

                                cep = c.ToString();

                                var po = cnode.GetScreenRectangle();
                                Mouse.Instance.Click(new System.Windows.Point(po.Value.X,po.Value.Y));
                                
                            }

                            if (ind == 5)
                            {
                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                var tit = cnode.GetScreenRectangle();

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                     
                                return cep;
                            }
                        }
                    }
                }
            }

            return cep;
        }

        public bool D_VerificaCidadeSP()
        {

            bool cidadeSP = false;
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("page tab: Processos", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                                var cnode = (AccessibleContextNode)noInterface;
                                int ind = cnode.GetIndexInParent();

                                if (ind == 16)
                                {
                                    var prop = cnode.GetProperties(PropertyOptions.AccessibleText);
                                    var acessible = (PropertyGroup)prop[0];
                                    var content = (PropertyGroup)acessible.Children[8];
                                    var c = content.Children[0].Value;

                                    if (c.ToString().ToUpper().Contains("SÃO PAULO") || c.ToString().ToUpper().Contains("SAO PAULO"))
                                    {
                                        cidadeSP = true;
                                    }

                                return cidadeSP;
                                }
                        }
                    }
                }
            }

            return cidadeSP;
        }

        public void D_BotaoEndereco_Foro()
        {

            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("page tab: Processos", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("Endereço"))
                            {
                                var cnode = (AccessibleContextNode)noInterface;

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                return;
                            }
                        }
                    }
                }
            }

        }

        public void D_LerOcorrenciasCadastroProcessoJuridico()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("page tab: Processos", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("Ocorrências Required"))
                            {
                                var cnode = (AccessibleContextNode)noInterface;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                var acessible = (PropertyGroup)prop[0];
                                var content = (PropertyGroup)acessible.Children[8];
                                var c = content.Children[0].Value;

                                string Nome = c.ToString();

                                if (c.ToString() == "No text contents")
                                {
                                    var contentPoint = (PropertyGroup)acessible.Children[7];
                                    var point = contentPoint.Children[0];
                                    var valuePoint = (AccessibleRectInfo)point.Value;
                                    Desktop.Instance.Mouse.Click(new System.Windows.Point(valuePoint.X + 10, valuePoint.Y + 2));

                                    int ind = cnode.GetIndexInParent();
                                    Thread.Sleep(1000);

                                    D_BotatoOcorrencia(ind);

                                    Thread.Sleep(1000);
                                    D_AcessarListaOcorrencias();

                                    D_BotaoSairProcessoJuridico();

                                    return;
                                }
                            }
                        }
                    }
                }

            }
        }

        public void D_BotatoOcorrencia(int index)
        {
            int indexButton = index + 10;

            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("page tab: Processos", n).GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren().ToList()[0].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("push button"))
                            {
                                var cnode = (AccessibleContextNode)noInterface;

                                int ind = cnode.GetIndexInParent();
                                
                                if(ind == indexButton)
                                {
                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                                    int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);

                                    return;
                                }
                            }
                        }
                    }
                }

            }
        }

        public void D_AcessarListaOcorrencias()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("Lista de Ocorrências", n).GetChildren().ToList()[0].GetChildren().ToList()[1].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("text: "))
                            {
                                var cnode = (AccessibleContextNode)noInterface;
                                var prop = cnode.GetProperties(PropertyOptions.AccessibleText);

                                var acessible = (PropertyGroup)prop[0];
                                var contentPoint = (PropertyGroup)acessible.Children[7];
                                var point = contentPoint.Children[0];
                                var valuePoint = (AccessibleRectInfo)point.Value;
                                Desktop.Instance.Mouse.Click(new System.Windows.Point(valuePoint.X, valuePoint.Y));

                                Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DELETE);
                                Keyboard.Instance.Enter("%173");

                                Thread.Sleep(1000);
                                D_BotaoProcurarOcorrencias();
                                //Thread.Sleep(1000);
                                D_BotaoOKOcorrencias();
                                D_BotaoSalvar();
                            }
                        }
                    }
                }

            }
        }

        public void D_BotaoProcurarForo()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("Lista de Foruns", n).GetChildren().ToList()[0].GetChildren().ToList()[3].GetChildren().ToList()[0].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("push button: Find"))
                            {
                                var cnode = (AccessibleContextNode)noInterface;

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                            }
                        }
                    }
                }
            }
        }


        public void D_BotaoProcurarOcorrencias()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("Lista de Ocorrências", n).GetChildren().ToList()[0].GetChildren().ToList()[3].GetChildren().ToList()[0].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("push button: Find"))
                            {
                                var cnode = (AccessibleContextNode)noInterface;

                                    AccessibleActions acc = null;
                                    _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                    AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                    {
                                        actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                        actionsCount = 1
                                    };

                                    ad1.actions[0].name = "Click";

                                    int falha1 = 0;

                                    _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                            }
                        }
                    }
                }
            }
        }

        public void D_BotaoOKListaForo()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("Lista de Foruns", n).GetChildren().ToList()[0].GetChildren().ToList()[3].GetChildren().ToList()[1].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("push button: OK"))
                            {
                                var cnode = (AccessibleContextNode)noInterface;

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                            }
                        }
                    }
                }
            }

        }

        public void D_BotaoOKOcorrencias()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noParentInterface = BuscaInterface("Lista de Ocorrências", n).GetChildren().ToList()[0].GetChildren().ToList()[3].GetChildren().ToList()[1].GetChildren();

                        foreach (var noInterface in noParentInterface)
                        {
                            var t = noInterface.GetTitle();

                            if (t.Contains("push button: OK"))
                            {
                                var cnode = (AccessibleContextNode)noInterface;

                                AccessibleActions acc = null;
                                _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                                AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                                {
                                    actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                                    actionsCount = 1
                                };

                                ad1.actions[0].name = "Click";

                                int falha1 = 0;

                                _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                            }
                        }
                    }
                }
            }
        }

        public void D_BotaoSalvar()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("Grava Informacoes", n);
                        var t = noInterface.GetTitle();

                        if (t.Contains("Grava Informacoes"))
                            {
                                var cnode = (AccessibleContextNode)noInterface;

                                Desktop.Instance.Mouse.Click(new System.Windows.Point(5 + 2, 51 + 2));

                        }
                    }
                }
            }
        }

        public void D_BotaoSairProcessoJuridico()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("Sai do programa", n);
                        var t = noInterface.GetTitle();

                        if (t.Contains("Sai do programa"))
                        {
                            var cnode = (AccessibleContextNode)noInterface;
                             Desktop.Instance.Mouse.Click(new System.Windows.Point(337+3, 51+4));
                        }
                    }
                }
            }
        }



        #endregion

        private AccessibleNode BuscaNo(AccessibleNode node, string titulo)
        {
            AccessibleNode retNode = null;

            foreach(var i in node.GetChildren())
            {
                var t = i.GetTitle();
                if (t.Contains(titulo))
                {
                    retNode = i;
                    return retNode;
                }

                if (i.GetChildren().Count() > 0)
                {
                    var ret = BuscaNo(i, titulo);

                    if (ret != null)
                    {
                        retNode = ret;

                        return retNode;
                    }
                }
            }

            return retNode;
        }
 
        private AccessibleNode BuscaInterface(string titulo, AccessibleNode node)
        {
            return BuscaNo(node, titulo);
        }

        #region Fluxo Custas ....

        public void C_Menu_CadastroProcessos()
        {
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("Cadastro De Processos", n);

                        var cnode = (AccessibleContextNode)noInterface;

                        AccessibleActions acc = null;
                        _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                        AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                        {
                            actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                            actionsCount = 1
                        };

                        ad1.actions[0].name = "Click";

                        int falha1 = 0;

                        var ret = _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                    }
                }
            }
        }

        public void C_MenuPendencias()
        {
            
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        bool novamente = true;
                        AccessibleNode noInterface = null;

                        while (novamente)
                        {
                            noInterface = BuscaInterface("page tab: Pendências", n);

                            if (noInterface != null)
                                novamente = false;
                        }


                        var cnode = (AccessibleContextNode)noInterface;
                        var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                        var acessible = (PropertyNode)prop[7];

                        var contentPoint = (AccessibleRectInfo)acessible.Value;
                        Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X + 20, contentPoint.Y + 10));
                    }
                }
            }
        }

        public void C_InserirOcorrencia(string ocorrencia)
        {
            //As ocorrências deverão ser separadas
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("As ocorrências deverão ser separadas", n);

                        var cnode = (AccessibleContextNode)noInterface;

                        var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                        var acessible = (PropertyNode)prop[7];

                        var contentPoint = (AccessibleRectInfo)acessible.Value;
                        Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X+10, contentPoint.Y+2));
                        Desktop.Instance.Enter(ocorrencia);

                    }
                }
            }
        }

        public void C_InserirGrupo(string grupo)
        {
            //combo box: Grupo
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("combo box: Grupo", n);

                        var cnode = (AccessibleContextNode)noInterface;
                        var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                        var acessible = (PropertyNode)prop[7];

                        var contentPoint = (AccessibleRectInfo)acessible.Value;
                        Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X+10, contentPoint.Y+2));
                        Desktop.Instance.Enter("G");
                        Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.RETURN);

                    }
                }
            }
        }

        public void C_InserirDataVencimento(string Data)
        {
            //text: Vencto até
            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("text: Vencto até", n);
                        var cnode = (AccessibleContextNode)noInterface;

                        var prop = cnode.GetProperties(PropertyOptions.AccessibleContextInfo);
                        var acessible = (PropertyNode)prop[7];

                        var contentPoint = (AccessibleRectInfo)acessible.Value;

                        DateTime dt = DateTime.Today.AddDays(5);
                        string dateFormat = string.Format("{0}/{1}/{2}",dt.Day.ToString("00"),dt.Month.ToString("00"), dt.Year.ToString("0000"));

                        Desktop.Instance.Mouse.Click(new System.Windows.Point(contentPoint.X+3, contentPoint.Y+2));


                        for (int i = 0; i < 11; i++) {
                           Keyboard.Instance.PressSpecialKey(TestStack.White.WindowsAPI.KeyboardInput.SpecialKeys.DELETE);
                        }

                        
                        Desktop.Instance.Enter(dateFormat);
                        
                    }
                }
            }

        }

        public void C_PesquisarOcorrencia()
        {
            //push button: Pesquisar

            var jvm = BuscarJanelaJava("Omni");

            if (jvm != null)
            {
                var nos = jvm.GetChildren();

                foreach (var n in nos)
                {
                    var no = n.GetTitle();

                    if (no.Contains("Omni"))
                    {
                        var noInterface = BuscaInterface("push button: Pesquisar", n);

                        var cnode = (AccessibleContextNode)noInterface;

                        AccessibleActions acc = null;
                        _accessBridge.Functions.GetAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, out acc);

                        AccessibleActionsToDo ad1 = new AccessibleActionsToDo()
                        {
                            actions = new AccessibleActionInfo[CodeGen.Interop.Constants.MAX_ACTION_INFO],
                            actionsCount = 1
                        };

                        ad1.actions[0].name = "Click";

                        int falha1 = 0;

                        var ret = _accessBridge.Functions.DoAccessibleActions(cnode.JvmId, cnode.AccessibleContextHandle, ref ad1, out falha1);
                    }
                }
            }
        }

        #endregion
    }
}
