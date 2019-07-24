using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsTeste.PlugIn;

namespace WindowsFormsTeste
{
    public partial class Form1 : Form
    {
        CobrancaPlugIn java = new CobrancaPlugIn();
        BizTalkPlugin biz = new BizTalkPlugin();
        OMNIPlugin omni = new OMNIPlugin();
        TRJPlugin trj = new TRJPlugin();
        public Form1()
        {
            InitializeComponent();

            java = new CobrancaPlugIn();
            java.InicializarDriver();

            biz = new BizTalkPlugin();
            trj = new TRJPlugin();

            //trj.BuildProcess()
            //.OpenWebProcess("http://www.tjsp.jus.br/CompetenciaTerritorial");

            //string foro = trj.RetornaForo_CEP("05616-060");

            //trj.Close();

            // Area para testes 
            java.LerArquivoImprimir("101891000036218");

            string teste = "";

            //java.FecharJanela("");


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

        private void AlteracaoFluxodigitalizacao()
        {
            java.JanelaDownloadIE();

            string nomearquivo = java.DownloadDocumento("100021000024016");

            if (nomearquivo.Length > 0)
            {
                LerArquivoGerarOcr(nomearquivo);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //if (MessageBox.Show("FLUXO ABERTURA - NOTIFICAÇÃO POSITIVA - DIGITAL ") == DialogResult.OK)
            //{
            //    java.AcessarMenuContratoParaMontagem();
            //    java.SelecionarTipoMontagem("N");
            //    java.SelecionarAdvogado("9");
            //    java.SelecionarDigital("S");
            //    java.SelecionarBotaoPesquisar();
            //    java.ClicardialogBoxAviso();
            //    java.LerTabelaBuscaApreensao("N","S");

            //}

            //if (MessageBox.Show("FLUXO ABERTURA - NOTIFICAÇÃO POSITIVA - FISICO ") == DialogResult.OK)
            //{
            //    java.AcessarMenuContratoParaMontagem();
            //    java.SelecionarTipoMontagem("N");
            //    java.SelecionarAdvogado("9");
            //    java.SelecionarDigital("N");
            //    java.SelecionarBotaoPesquisar();
            //    java.ClicardialogBoxAviso();
            //    java.LerTabelaBuscaApreensao("N","F");

            //}

            if (MessageBox.Show("FLUXO ABERTURA - PROTESTO - DIGITAL ", "FLUXOS", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                java.AcessarMenuContratoParaMontagem();
                java.SelecionarTipoMontagem("P");
                java.SelecionarAdvogado("9");
                java.SelecionarDigital("S");
                java.SelecionarBotaoPesquisar();
                java.ClicardialogBoxAviso();
                java.LerTabelaBuscaApreensao("P", "S");
                java.ScroolTabelaBuscaApreensao();
                java.LerTabelaBuscaApreensao("P", "S");
                java.ScroolTabelaBuscaApreensao();
                java.LerTabelaBuscaApreensao("P", "S");

            }

            if (MessageBox.Show("FLUXO ABERTURA - PROTESTO - FISICO ", "FLUXOS", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
            
                java.AcessarMenuContratoParaMontagem();
                java.SelecionarTipoMontagem("P");
                java.SelecionarAdvogado("9");
                java.SelecionarDigital("N");
                java.SelecionarBotaoPesquisar();
                java.ClicardialogBoxAviso();
                java.LerTabelaBuscaApreensao("P", "F");
                //java.ScroolTabelaBuscaApreensao();
                //java.LerTabelaBuscaApreensao("P", "F");

            }

            //if (MessageBox.Show("FLUXO ABERTURA - ASSESSORIA ") == DialogResult.OK)
            //{
            //    java.AcessarMenuContratoParaMontagem();
            //    java.SelecionarTipoMontagem("P");

            //    java.SelecionarAdvogado("0");
                
            //    java.SelecionarDigital("N");
            //    java.SelecionarBotaoPesquisar();
            //    java.ClicardialogBoxAviso();
            //    java.LerTabelaBuscaApreensao("A","S");

            //}

            //java.LerTabelaBuscaApreensao();
            //java.ClicarRetornarTelaPrincipal();

            //java.AbrirDocumento_DistribuicaoProcesso(186);
            //java.ValidarCheckDocumento_DistribuicaoProcesso(130);
            //java.ValidarDocumentoDigital_DistribuicaoProcesso(88);

            //string doc = "100324000053017";

            MessageBox.Show("Fluxos Abertura Finalizados");

        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("Processo Abrir Sistema OMNI , certifique se de que não esteja nenhuma instancia aberta ") == DialogResult.OK)
            {
                // Abrir OMNI 
                omni.BuildProcess()
                   .OpenWebProcess("http://oracle.omni.com.br:8888/forms/frmservlet?config=prod")
                    .RealizarLogin();

                
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("FLUXO DISTRIBUIÇÃO - REMOÇÂO 173 ", "FLUXOS OMNI", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                java.AcessarMenuContratoParaMontagem();
                java.D_Menu_MontageRemontagem();
                java.D_SelecionarTipoMontagem("M");
                java.D_SelecionarAdvogado("968");
                java.D_SelecionarTipomontagemNotificacao("N");
                java.D_SelecionarBotaoPesquisar();
                java.ClicardialogBoxAviso();
                java.D_LerTabelaBuscaApreensao("");
            }

            if (MessageBox.Show("FLUXO DISTRIBUIÇÃO - CADASTRO FORO ","FLUXOS OMNI", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {

                    java.AcessarMenuContratoParaMontagem();
                    java.D_Menu_MontageRemontagem();
                    java.D_SelecionarTipoMontagem("M");
                    java.D_SelecionarAdvogado("968");
                    java.D_SelecionarTipomontagemNotificacao("N");
                    java.D_SelecionarBotaoPesquisar();
                    java.ClicardialogBoxAviso();
                    java.D_LerTabelaBuscaApreensao_CadastroForo("");
            }

            if (MessageBox.Show("FLUXO DISTRIBUIÇÃO - LIBERAR CUSTAS ", "FLUXOS OMNI", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {

                java.AcessarMenuContratoParaMontagem();
                java.D_Menu_MontageRemontagem();
                java.D_SelecionarTipoMontagem("M");
                java.D_SelecionarAdvogado("968");
                java.D_SelecionarTipomontagemNotificacao("N");
                java.D_SelecionarBotaoPesquisar();
                java.ClicardialogBoxAviso();
                java.D_LerTabelaBuscaApreensao_CadastroForo("");
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("FLUXO CUSTAS -  ") == DialogResult.OK)
            {
                java.C_Menu_CadastroProcessos();
                java.C_MenuPendencias();
                java.C_InserirOcorrencia("352");
                java.C_InserirGrupo("G");
                java.C_InserirDataVencimento("");
                java.C_PesquisarOcorrencia();
            }

        }
    }
}
