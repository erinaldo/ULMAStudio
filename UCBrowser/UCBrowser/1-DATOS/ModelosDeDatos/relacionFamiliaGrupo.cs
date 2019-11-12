using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UCBrowser
{
    public class relacionFamiliaGrupo
    {
        public string idFamilia;
        public string idGrupoAlQuePertenece;

        private relacionFamiliaGrupo() { }

        public relacionFamiliaGrupo(string idFamilia, string idGrupoAlQuePertenece)
        {
            this.idFamilia = idFamilia;
            this.idGrupoAlQuePertenece = idGrupoAlQuePertenece;
        }

        public override bool Equals(object otro)
        {
            if (otro.GetType() == this.GetType())
            {
                relacionFamiliaGrupo otraRelacion = (relacionFamiliaGrupo)otro;
                if (this.idFamilia.Equals(otraRelacion.idFamilia)
                 && this.idGrupoAlQuePertenece.Equals(otraRelacion.idGrupoAlQuePertenece))
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
            return this.idFamilia.GetHashCode() + this.idGrupoAlQuePertenece.GetHashCode();
        }


    }
}
