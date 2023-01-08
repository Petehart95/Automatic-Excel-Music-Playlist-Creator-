using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using SpreadsheetLight;

namespace AutomaticExcelMusicPlaylistCollator
{
    class FileProcessor 
    {
        public static int trackctr;
        public static SLDocument doc;
        public string path;
        public int totalFiles;

        static FileProcessor()
        {
            trackctr = 0;
            doc = new SLDocument();
        }

        public void ProcessDirectory(string targetDirectory)
        {
            bool newArtist;
            string lastFolderName = ""; string currentFolderName;

            string[] fileEntries = Directory.GetFiles(targetDirectory);


            foreach (string fileName in fileEntries)
            {

                currentFolderName = Path.GetFileName(Path.GetDirectoryName(fileName));
                newArtist = false;
                if (Path.GetExtension(fileName).ToUpper() == ".MP3" || Path.GetExtension(fileName).ToUpper() == ".WAV")
                {
                    trackctr = trackctr + 1;

                    if (lastFolderName != currentFolderName)
                    {
                        lastFolderName = currentFolderName; newArtist = true;
                    }
                    ProcessFile(fileName, newArtist);

                }
            }

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory);
            }
        }
        public int GetDirectoryLength(string path)
        {
            return Directory.GetFiles(path, "*", SearchOption.AllDirectories).Length;
        }

        public void ProcessFile(string path, bool flag)
        {
            TagLib.File tagFile = TagLib.File.Create(path);
            string title = tagFile.Tag.Title;
            string lastFolderName = Path.GetFileName(Path.GetDirectoryName(path));
            string artistCell; string trackCell;

            if (flag == true)
            {
                artistCell = 'A' + trackctr.ToString();
                doc.SetCellValue(artistCell, lastFolderName);

                Console.WriteLine('\n' + lastFolderName);
            }

            trackCell = 'B' + trackctr.ToString();
            doc.SetCellValue(trackCell, title);

            Console.Write(trackctr + " - " + title);
            Console.WriteLine();
        }

        public bool CreateNewFile(string pFileName, string pSheetName)
        {
            doc.RenameWorksheet("Sheet1", pSheetName);
            doc.SaveAs(pFileName);

            return true;
        }

        public bool AddNewSheet(string pFileName, string pSheetName)
        {
            if (!(doc.GetSheetNames(false).TrueForAll((sheetName) => sheetName.ToLower() == pSheetName.ToLower())))
            {
                doc.AddWorksheet(pSheetName);
                doc.Save();

                return true;
            }
            else
            {
                return false;
            }

        }
        public void SaveFile(string pFileName)
        {

            Cursor.Current = Cursors.WaitCursor;

            doc.SaveAs(pFileName);

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Set cursor as default arrow
            Cursor.Current = Cursors.Default;
        }
    }
}

