using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace DocToDocxConverter
{
    public static class RecycleBin
    {
        public static void DeleteFile(string file)
        {
            Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(file, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
            if (File.Exists(file))
            {
                throw new Exception("Assert failed. Failed to delete the file while no errors were thrown. This should not happen.");
            }
        }
    }
}
