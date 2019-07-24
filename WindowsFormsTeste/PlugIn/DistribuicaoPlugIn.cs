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

namespace WindowsFormsTeste.PlugIn
{
    public class DistribuicaoPlugIn
    {
        AccessBridge _accessBridge = new AccessBridge();
        private Application _application;
        public Application Application { get { return _application; } set { _application = value; } }
        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        BizTalkPlugin biz = new BizTalkPlugin();

        private Window _window;
        public Window Window { get { return _window; } set { _window = value; } }

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

    }
}
