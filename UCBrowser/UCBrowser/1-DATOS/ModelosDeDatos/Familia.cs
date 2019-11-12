using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCBrowser
{

    public class Familia : ElementoDeLaBiblioteca
    {
        public string descripcion { get; set; }
        public string tooltip { get; set; }

        public readonly string nombreArchivo;
        //private string _nombreArchivo;
        //public string nombreArchivo { get { return _nombreArchivo; } }
        public readonly bool esDinamica;
        public readonly bool esConjunto;
        public readonly bool esAnnotationSymbol;

        public Familia(string id, string nombre, string nombreArchivo,
                       bool esDinamica, bool esConjunto, bool esAnnotationSymbol)
            : base(id, nombre)
        {
            this.descripcion = this.nombre;
            this.nombreArchivo = nombreArchivo;
            this.esDinamica = esDinamica;
            this.esConjunto = esConjunto;
            this.esAnnotationSymbol = esAnnotationSymbol;
            this.tooltip = this.nombre + Environment.NewLine + Environment.NewLine + this.nombreArchivo;
        }

        public override bool Equals(object obj)
        {
            if (obj.GetType() == this.GetType())
            {
                Familia otraFamilia = (Familia)obj;
                if (this.nombreArchivo.Equals(otraFamilia.nombreArchivo))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode()
        {
            return this.nombreArchivo.GetHashCode();
        }

    }


    public class ComparadorDeFamilias : IEqualityComparer<Familia>
    {
        public bool Equals(Familia A, Familia B)
        {
            if (A.nombreArchivo.Equals(B.nombreArchivo))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public int GetHashCode(Familia familia)
        {
            return familia.nombreArchivo.GetHashCode();
        }
    }



}
