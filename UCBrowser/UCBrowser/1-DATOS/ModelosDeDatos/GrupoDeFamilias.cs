using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCBrowser
{
    public class GrupoDeFamilias : ElementoDeLaBiblioteca
    {
        public readonly string lineaALaQuePertenece;
        public readonly string nombreCorto;
        public string descripcion { get; set; }

        public GrupoDeFamilias(string id, string nombre, string lineaALaQuePertenece, string nombreCorto)
            : base(id, nombre)
        {
            this.lineaALaQuePertenece = lineaALaQuePertenece;
            this.nombreCorto = nombreCorto;
            this.descripcion = nombre;
        }
    }
}
