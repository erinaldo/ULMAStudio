using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCBrowser
{
    public class LineaDeProducto : ElementoDeLaBiblioteca
    {
        public LineaDeProducto(string id, string nombre)
            : base(id, nombre)
        {
            switch (id)
            {
                case "AND":
                    thumbnail = Properties.Resources.AND;
                    break;
                case "EHL":
                    thumbnail = Properties.Resources.EHL;
                    break;
                case "EHP":
                    thumbnail = Properties.Resources.EHP;
                    break;
                case "EV":
                    thumbnail = Properties.Resources.EV;
                    break;
                case "STR":
                    thumbnail = Properties.Resources.STR;
                    break;
                case "COM":
                    thumbnail = Properties.Resources.COM;
                    break;
                case "ANNOTATIONS":
                    thumbnail = Properties.Resources.ANNOTATIONS;
                    break;
            }
        }

    }
}
