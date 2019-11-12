using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCBrowser
{
    public abstract class ElementoDeLaBiblioteca
    {
        private string _id;
        public string id { get { return _id; } }

        private string _nombre;
        public string nombre { get { return _nombre; } }

        private System.Drawing.Bitmap _thumbnail;
        public System.Drawing.Bitmap thumbnail
        { 
            get { return _thumbnail; }
            set
            {
                if (value == null){

                }
                else
                {
                    _thumbnail = value;
                }

            }
        }


        public ElementoDeLaBiblioteca(string id, string nombre)
        {
            this._id = id;
            this._nombre = nombre;
            this._thumbnail = Properties.Resources._default;
        }

        public override string ToString()
        {
            return "[" + id + "] " + nombre;
        }

    }
}
