using Aspose.Cells;
using System;
using System.IO;

namespace imageMani
{


    class Program
    {
        static void Main(string[] args)
        {
            // license read
            ReadLicence();
            try
            {
                //"@"E:\Codes\C# Workspace\imageMani\imageMani\data\PVC-Cartographie-201218_171118.xlsx"
                string inputFilePath = args[0];
                if (!File.Exists(inputFilePath))
                    throw new Exception(Path.GetFileNameWithoutExtension(inputFilePath) + "Not Found");
                new TreatExcel(inputFilePath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            // directory





        }

        static void ReadLicence()
        {
            Console.WriteLine("Checking License");
            // Instantiate the License object
            License lic = new License();
            // Set the License stream
            lic.SetLicense("Aspose.Total.lic");

        }
    }
}
