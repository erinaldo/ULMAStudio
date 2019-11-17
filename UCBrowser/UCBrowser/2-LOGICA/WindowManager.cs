using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;


namespace UCBrowser
{
    public static class WindowManager
    {
        public static event EventHandler<EventArgs> ClosingWindows;

        public static void CloseWindows()
        {
            var c = ClosingWindows;
            if (c != null)
            {
                c(null, EventArgs.Empty);
            }
        }
    }
}
