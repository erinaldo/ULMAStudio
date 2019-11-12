using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCBrowser
{
    public class Filial
    {
        private int _id;
        public int id { get { return _id; } }
        private string _nombre;
        public string nombre { get { return _nombre; } }
        private string _idioma;
        public string idioma { get { return _idioma; } }

        public Filial(int id, string nombre, string idioma)
        {
            this._id = id;
            this._nombre = nombre;
            this._idioma = idioma;
        }

        public override string ToString()
        {
            return "[" + id + "] " + nombre;
        }
    }

}
