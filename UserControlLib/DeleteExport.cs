using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserControlLib
{
    public class DeleteExport
    {
        public static void clearexport()
        {
            string rootFolder = AppDomain.CurrentDomain.BaseDirectory;
            string exportFolder = System.IO.Path.Combine(rootFolder, "export");

            if (Directory.Exists(exportFolder))
            {
                Directory.Delete(exportFolder, true);
            }
        }

    }
}
