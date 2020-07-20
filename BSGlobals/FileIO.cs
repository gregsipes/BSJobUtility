using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BSGlobals
{
   public class FileIO
    {
        public static void CheckCreateDirectory(string filePath)
        {
            CheckCreateDirectory(filePath, false);
        }

        public static void CheckCreateDirectory(string filePath, bool containsFileName)
        {
            string directory = "";
            if (containsFileName)
                directory = Path.GetDirectoryName(filePath);
            else
                directory = filePath;


            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
        }


        public static List<string> GetFiles(string sourceDirectory, Regex reg)
        {
            // validate existence of directory
            CheckCreateDirectory(sourceDirectory);

            return Directory.GetFiles(sourceDirectory)
                .Where(f => ((reg == null) ? true : reg.IsMatch(Path.GetFileName(f))))
                .ToList();
        }

    }
}
