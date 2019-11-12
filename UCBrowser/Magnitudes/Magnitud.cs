using System;
using System.Windows;
using System.IO;
using System.Xml;
using System.Data;


namespace Magnitudes
{
    public class Magnitud : IFormattable , IEquatable<Object> , IComparable<Magnitud>
    {
        public double valor;

        private string _unidaddemedida;
        public string unidaddemedida
         {
            get { return _unidaddemedida; } 
            set { if (string.IsNullOrWhiteSpace(value)) {
                _unidaddemedida = DESCONOCIDA;
                  } 
                  else {
                      _unidaddemedida = value;
                  }
                }
         }
        public const string DESCONOCIDA = "Desconocida-Unknow";

        public Magnitud()
        {
            this.valor = double.NaN;
            this.unidaddemedida = DESCONOCIDA;
        }

        public Magnitud(double valor, string unidadDeMedida)
        {
            this.valor = valor;
            this.unidaddemedida = unidadDeMedida;
        }

        public Magnitud(double valor)
        {
            this.valor = valor;
            this.unidaddemedida = DESCONOCIDA;
        }

        public Magnitud(Magnitud copiarDe)
        {
            this.valor = copiarDe.valor;
            this.unidaddemedida = copiarDe.unidaddemedida;
        }


        public Magnitud ConvertirmeA(string nuevaUnidadDeMedida)
        {
            Magnitud conversion = ConversorDeMagnitudes.Convertir(magnitudOrigen:this, unidadDestino:nuevaUnidadDeMedida);
            this.valor = conversion.valor;
            this.unidaddemedida = conversion.unidaddemedida;
            return this;
        }

        public Magnitud getCopiaConvertidaA(string nuevaUnidadDeMedida)
        {
            return new Magnitud(ConversorDeMagnitudes.Convertir(magnitudOrigen:this, unidadDestino:nuevaUnidadDeMedida));
        }


        public override string ToString()
        {
            return this.valor.ToString() + " " + this.unidaddemedida;
        }

        public string ToString(string formatoDeNumero)
        {
            return this.valor.ToString(formatoDeNumero) + " " + this.unidaddemedida;
        }

        public string ToString(string formatoDeNumero, IFormatProvider provider)
        {
            return this.valor.ToString(formatoDeNumero, provider) + " " + this.unidaddemedida;
        }

        public string ToStringConCulturaNeutra()
        {
            return this.valor.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + this.unidaddemedida;
        }


        public XmlNode ToXML(XmlDocument docXML, string etiqueta)
        {
            XmlNode nodo = docXML.CreateElement(etiqueta);
            nodo.InnerText = this.valor.ToString(System.Globalization.CultureInfo.InvariantCulture);
            XmlAttribute atributo = docXML.CreateAttribute("unit");
            atributo.InnerText = this.unidaddemedida;
            nodo.Attributes.Append(atributo);
            return nodo;
        }

        public void FromXML(XmlNode nodo, string etiqueta)
        {
            if (nodo.Name.Equals(etiqueta))
            {
                this.unidaddemedida = nodo.Attributes.Item(0).InnerText;
                double.TryParse(s: nodo.InnerText,
                                style: System.Globalization.NumberStyles.Any, provider: System.Globalization.CultureInfo.InvariantCulture,
                                result: out this.valor);
            }
            else
            {
                throw new ArgumentException(nodo.Name + " != " + etiqueta);
            }
        }
        public Magnitud(XmlNode nodo)
        {
            this.unidaddemedida = nodo.Attributes.Item(0).InnerText;
            double.TryParse(s:nodo.InnerText, 
                            style:System.Globalization.NumberStyles.Any, provider:System.Globalization.CultureInfo.InvariantCulture, 
                            result: out this.valor);
        }




        public override bool Equals(object otroObjeto)
        {
            if (otroObjeto == null)
            {
                return false;
            }
            if (otroObjeto.GetType().Equals(this.GetType()))
            {
                Magnitud otraMagnitud = (Magnitud) otroObjeto;
                if (otraMagnitud.getCopiaConvertidaA(this.unidaddemedida).valor.Equals(this.valor))
                {
                    return true;
                }
            }
            return false;
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }


        public int CompareTo(Magnitud otraMagnitud)
        {
            if (this.Equals(otraMagnitud))
            {
                return 0;
            }
            else if (this.getCopiaConvertidaA(otraMagnitud.unidaddemedida).valor > this.valor)
            {
                return 1;
            }
            else
            {
                return -1;
            }
        }


        public static Magnitud operator +(Magnitud unaMagnitud, Magnitud otraMagnitud)
        {
            return new Magnitud(valor: unaMagnitud.valor + otraMagnitud.getCopiaConvertidaA(unaMagnitud.unidaddemedida).valor,
                                unidadDeMedida: unaMagnitud.unidaddemedida);
        }

        public static Magnitud operator -(Magnitud unaMagnitud, Magnitud otraMagnitud)
        {
            return new Magnitud(valor: unaMagnitud.valor - otraMagnitud.getCopiaConvertidaA(unaMagnitud.unidaddemedida).valor,
                                unidadDeMedida: unaMagnitud.unidaddemedida);
        }

        public static Magnitud operator *(Magnitud unaMagnitud, Magnitud otraMagnitud)
        {
            Magnitud resultado;
            Magnitud otra = otraMagnitud.getCopiaConvertidaA(unaMagnitud.unidaddemedida);
            if (otra.unidaddemedida == Magnitud.DESCONOCIDA)
            {
                resultado = new Magnitud(valor: unaMagnitud.valor * otraMagnitud.valor,
                                         unidadDeMedida: unaMagnitud.unidaddemedida + "*" + otraMagnitud.unidaddemedida);
            }
            else
            {
                resultado = new Magnitud(valor: unaMagnitud.valor * otra.valor,
                                         unidadDeMedida: unaMagnitud.unidaddemedida + "2");
            }
            //string _pendiente_NoSeHanContempladoTodosLosCasosPosiblesEnLasUnidadesDeMedidaDelResultado;
            return resultado;
        }

        public static Magnitud operator /(Magnitud unaMagnitud, Magnitud otraMagnitud)
        {
            Magnitud resultado;
            Magnitud otra = otraMagnitud.getCopiaConvertidaA(unaMagnitud.unidaddemedida);
            if (otra.unidaddemedida == Magnitud.DESCONOCIDA)
            {
                resultado = new Magnitud(valor: unaMagnitud.valor / otraMagnitud.valor,
                                         unidadDeMedida: unaMagnitud.unidaddemedida + "/" + otraMagnitud.unidaddemedida);
            }
            else
            {
                resultado = new Magnitud(valor: unaMagnitud.valor / otra.valor,
                                         unidadDeMedida: "");
            }
            //string _pendiente_NoSeHanContempladoTodosLosCasosPosiblesEnLasUnidadesDeMedidaDelResultado;
            return resultado;
        }

        public static Magnitud operator *(Magnitud unaMagnitud, double multiplicador)
        {
            return new Magnitud(valor: unaMagnitud.valor * multiplicador,
                                unidadDeMedida: unaMagnitud.unidaddemedida);
        }



    }




    class ConversorDeMagnitudes
    {
        // unidades de base.
        //==================
        public  const string UNIDADPARA_LONGITUD = "m"; 
        public  const string UNIDADPARA_MASA = "kg";
        public  const string UNIDADPARA_TIEMPO = "s";  
        public  const string UNIDADPARA_ANGULO = "rad";
        public  const string UNIDADPARA_TEMPERATURA = "C"; 

        // principales unidades derivadas.
        //================================
        public  const string UNIDADPARA_AREA = UNIDADPARA_LONGITUD + "2";  // m2
        public  const string UNIDADPARA_VOLUMEN = UNIDADPARA_LONGITUD + "3";  // m3
        public  const string UNIDADPARA_INERCIA = UNIDADPARA_LONGITUD + "4";  // m4

        public  const string UNIDADPARA_DENSIDAD = UNIDADPARA_MASA + "/" + UNIDADPARA_VOLUMEN; 

        //public static const string UNIDADPARA_MASA = "N"; // _pendiente_ de ver que hacemos con el kg-masa y el kg-peso(fuerza)
        public  const string UNIDADPARA_FUERZA = "N";  // N = kg*m/s2
        public  const string UNIDADPARA_RIGIDEZ = UNIDADPARA_FUERZA + "/" + UNIDADPARA_LONGITUD;  // N/m
        public  const string UNIDADPARA_PRESION = UNIDADPARA_FUERZA + "/" + UNIDADPARA_AREA;  // N/m2
        public  const string UNIDADPARA_TENSION = UNIDADPARA_PRESION;
        public  const string UNIDADPARA_MOMENTO = UNIDADPARA_FUERZA + "x" + UNIDADPARA_LONGITUD;  // Nxm
        public  const string UNIDADPARA_PARTORSOR = UNIDADPARA_MOMENTO;

        public  const string UNIDADPARA_VELOCIDAD = UNIDADPARA_LONGITUD + "/" + UNIDADPARA_TIEMPO;  // m/s  
        public  const string UNIDADPARA_ACELERACION = UNIDADPARA_LONGITUD + "/" + UNIDADPARA_TIEMPO + "2";  // m/s2  
        public  const string UNIDADPARA_ACELERACIONANGULAR = UNIDADPARA_ANGULO + "/" + UNIDADPARA_TIEMPO + "2";  // rad/s2  
        public  const string UNIDADPARA_FRECUENCIA = "/" + UNIDADPARA_TIEMPO;  //  Hz = 1/s  


        static ConversorDeMagnitudes m_conversor;
        static DataTable m_factoresDeConversion;

        private static ConversorDeMagnitudes getConversor()
        {
            if (m_conversor == null)
            {
                m_conversor = new ConversorDeMagnitudes();
            }
            return m_conversor;
        }

        private ConversorDeMagnitudes()
        {
            m_factoresDeConversion = new DataTable("FactoresDeConversion");
            m_factoresDeConversion.Columns.Add("unidadOrigen", System.Type.GetType("System.String"));
            m_factoresDeConversion.Columns.Add("unidadDestino", System.Type.GetType("System.String"));
            m_factoresDeConversion.Columns.Add("factorDeConversion", System.Type.GetType("System.Double"));

            string archivoDeFactoresDeConversion = Path.GetDirectoryName(System.Reflection.Assembly.GetCallingAssembly().Location) 
                                                   + Path.DirectorySeparatorChar + "conversionFactors.xml";
            if(File.Exists(archivoDeFactoresDeConversion)) {
                XmlDocument docXML = new XmlDocument();
                docXML.Load(archivoDeFactoresDeConversion);
                XmlNodeList factores = docXML.SelectNodes("/conversionFactors/conversion");
                foreach(XmlNode nConversion in factores) {
                    try 
	                {
                        DataRow factor = m_factoresDeConversion.NewRow();
                        factor["unidadOrigen"] = nConversion.Attributes.GetNamedItem("original").InnerText;
                        factor["unidadDestino"] = nConversion.Attributes.GetNamedItem("target").InnerText;
                        factor["factorDeConversion"] = Double.Parse(nConversion.InnerText, System.Globalization.CultureInfo.InvariantCulture);	        
		                m_factoresDeConversion.Rows.Add(factor);
	                }
	                catch (Exception)
	                {
		                MessageBox.Show(messageBoxText: Properties.Resources.msgErrorLeyendoArchivoDeFactoresDeConversion
                                                       + Environment.NewLine + Environment.NewLine 
                                                       + nConversion.OuterXml
                                                       + Environment.NewLine + Environment.NewLine 
                                                       + archivoDeFactoresDeConversion,
                                        caption: "", button: MessageBoxButton.OK, icon: MessageBoxImage.Warning);
	                }
                }
            }
            else {
               MessageBox.Show(messageBoxText: Properties.Resources.msgNoExisteElArchivo
                                                + Environment.NewLine + Environment.NewLine 
                                                + archivoDeFactoresDeConversion,
                                caption: "", button: MessageBoxButton.OK, icon: MessageBoxImage.Error);
            }
        }

        public static Magnitud Convertir(Magnitud magnitudOrigen, string unidadDestino)
        {
            getConversor();
            if (unidadDestino == magnitudOrigen.unidaddemedida || unidadDestino == Magnitud.DESCONOCIDA
                || magnitudOrigen.unidaddemedida == Magnitud.DESCONOCIDA || double.IsNaN(magnitudOrigen.valor))
            {
                return magnitudOrigen;
            }
            else
            {
                DataRow[] factores = m_factoresDeConversion.Select("unidadOrigen='" + magnitudOrigen.unidaddemedida + "'"
                                                                   + "AND unidadDestino='" + unidadDestino + "'");
                if (factores.Length == 1)
                {
                    return new Magnitud(magnitudOrigen.valor * factores[0].Field<double>("factorDeConversion"), unidadDestino);
                }
                else
                {
                    DataRow[] factoresInversos = m_factoresDeConversion.Select("unidadOrigen='" + unidadDestino + "'"
                                                                               + "AND unidadDestino='" + magnitudOrigen.unidaddemedida + "'");
                    if (factoresInversos.Length == 1)
                    {
                        return new Magnitud(magnitudOrigen.valor / factoresInversos[0].Field<double>("factorDeConversion"), unidadDestino);
                    }
                    else
                    {
                        MessageBox.Show(messageBoxText: Properties.Resources.msgFactorDeConversionNoEncontrado + Environment.NewLine + Environment.NewLine
                                                       + magnitudOrigen.ToString() + " = ? " + unidadDestino + Environment.NewLine,
                                        caption: "", button: MessageBoxButton.OK, icon: MessageBoxImage.Warning);
                        return new Magnitud(double.NaN, Magnitud.DESCONOCIDA);
                    }
                }
            }
        }

    }


}
