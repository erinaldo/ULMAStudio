﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Magnitudes.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    public class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("Magnitudes.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to PROBLEMA: se ha producido una división por cero al intentar convertir entre las unidades de medida indicadas..
        /// </summary>
        public static string msgDivisionPorCeroEnModuloConversorUnidades {
            get {
                return ResourceManager.GetString("msgDivisionPorCeroEnModuloConversorUnidades", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to ERROR GRAVE: hay algún problema con el archivo de factores de conversión entre distintas unidades de medida..
        /// </summary>
        public static string msgErrorLeyendoArchivoDeFactoresDeConversion {
            get {
                return ResourceManager.GetString("msgErrorLeyendoArchivoDeFactoresDeConversion", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to No se ha podido convertir este valor, por no disponer del factor de conversión correspondiente entre las unidades de medida indicadas..
        /// </summary>
        public static string msgFactorDeConversionNoEncontrado {
            get {
                return ResourceManager.GetString("msgFactorDeConversionNoEncontrado", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to ERROR GRAVE: no se ha podido encontrar el archivo indicado..
        /// </summary>
        public static string msgNoExisteElArchivo {
            get {
                return ResourceManager.GetString("msgNoExisteElArchivo", resourceCulture);
            }
        }
    }
}
