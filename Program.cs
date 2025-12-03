using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SBAAddon
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //clstesting ts = new clstesting();
           // clsnewforsba gt = new clsnewforsba();
              ClsAddon tst = new ClsAddon();
            Application.Run();
        }
    }
}
