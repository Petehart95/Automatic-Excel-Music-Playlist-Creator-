using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using SpreadsheetLight;

namespace AutomaticExcelMusicPlaylistCollator
{
    public class MusicCollator
    {
        [STAThread]

        public static void Main()
        {
            string outDir, fileName;
            FileProcessor fileProcessor = new FileProcessor();
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            Console.WriteLine("////////////////////////////////////////////////////");
            Console.WriteLine("\tAutomatic Music Playlist Spreadsheet");
            Console.WriteLine("////////////////////////////////////////////////////\n");

            Console.WriteLine("Press any key to begin...");
            Console.ReadKey();

            Console.WriteLine("Please select the main music folder...");

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                fileProcessor.path = fbd.SelectedPath;
                fileProcessor.totalFiles = fileProcessor.GetDirectoryLength(fileProcessor.path);
                fileName = "myMusicPlaylist.xlsx";

                outDir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.Parent.Parent.FullName;
                outDir = outDir + "\\" + fileName;

                Console.WriteLine("\nWill be Looking at files in this folder:");
                Console.WriteLine(fileProcessor.path);
                Console.WriteLine("Total Files Found Approximately (including non-music files): {0}", fileProcessor.totalFiles.ToString());

                Console.WriteLine();
                Console.WriteLine("This program will create an excel file based on all the music files in this folder.");
                Console.WriteLine();

                Console.WriteLine("Press any key to start...");
                Console.ReadLine();

                if (System.IO.Directory.Exists(fileProcessor.path))
                {
                    fileProcessor.ProcessDirectory(fileProcessor.path);
                }
                else
                {
                    Console.WriteLine("{0} is not a valid file or directory.", fileProcessor.path);
                    Console.ReadLine();
                }

                Console.WriteLine();
                Console.WriteLine("All Files Found! Now creating an Excel file...");
                Console.WriteLine();
                fileProcessor.SaveFile(outDir);

                Console.WriteLine("DONE! Spreadsheet has now been made! The spreadsheet file should be here: ");
                Console.WriteLine();
                Console.WriteLine(outDir);
                Console.ReadLine();
            }
        }
    }
}
