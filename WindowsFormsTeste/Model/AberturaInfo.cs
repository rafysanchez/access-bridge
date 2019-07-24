using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsTeste.Model
{

    public class OCRDocumento
    {
        public string TipoDocumento { get; set; }
        public string OCR { get; set; }
        public string NomeOCR { get; set; }
        public string Endereco { get; set; }
        public string Rua { get; set; }
        public string Bairro { get; set; }
        public string Cidade { get; set; }
        public string Estado { get; set; }
        public string CEP { get; set; }
    }

    public class AberturaInfo
    {
        public string Nome { get; set; }
        public string Contrato { get; set; }
        public string Tipo { get; set; }
        public bool Digital { get; set; }
        public StringBuilder Eventos { get; set; }
        public bool FluxoExcecao { get; set; }
        List<OCRDocumento> documentos { get; set; }
    }
}
