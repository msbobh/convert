using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using word = Microsoft.Office.Interop.Word;

namespace convert
{
    class Program
    {
        static void Main(string[] args)
        {
            word.Document doc = null;
            // Get the current directory, create a subdirectory for the output files
            string path = Directory.GetCurrentDirectory();
            string outputdir = System.IO.Path.Combine(path, "converted_files\\");
            System.IO.Directory.CreateDirectory(outputdir);
            
            // Enumerate the files to be converted
            string[] docfiles = Directory.GetFiles(path, "*.doc");

            word.Application wordapp = new word.Application();
            wordapp.Visible = false;
            string save;
            int count = 0;
            
            Console.WriteLine("{0} Files found", docfiles.Count());
            foreach (var file in docfiles)
            {
                Console.WriteLine("Opening file: {0}", file);
                object temp = file;
                try
                {
                    doc = wordapp.Documents.Open(ref temp);
                    // Change the suffix to docx, extract thh filename and append to new path w/ new subdirectory
                    save = file.Replace("doc", "docx");
                    save = Path.GetFileName(save);
                    save = Path.Combine(outputdir + save);

                    Console.WriteLine("Writing file to: {0}", save);
                    doc.SaveAs2(save, word.WdSaveFormat.wdFormatXMLDocument);
                    count++;
                }
                catch (Exception e)
                {
                    if (file == null) doc.Close();
                    if (wordapp == null) wordapp.Quit();
                    Console.WriteLine("Exception {0}", e, "Error writing file");
                    break;
                }

            }
            if (doc != null)
                doc.Close();
            wordapp.Quit();
            Console.WriteLine("{0} files processed", count);
        }
    }
    
}
