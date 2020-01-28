using MDLInputFileSetting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MDL
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            DevExpress.Skins.SkinManager.Default.RegisterAssembly(typeof(DevExpress.UserSkins.SkinProject2).Assembly);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
               Application.Run(new Login());
            }
            catch (Exception ex)
            {

            }
        }
    }
}
