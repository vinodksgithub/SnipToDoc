using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // Set your folder path here
        string searchFolderPath = @"C:\Users\91974\source\repos\SnipToDoc\SnipToDocuVer1.61\";

        // Output file path (Notepad can open any .txt file)
        string outputFilePath = Path.Combine(searchFolderPath, "CollectedFiles.txt");

        using (StreamWriter writer = new StreamWriter(outputFilePath))
        {
            // Collect App.config if present
            string appConfigPath = Path.Combine(searchFolderPath, "App.config");
            if (File.Exists(appConfigPath))
            {
                writer.WriteLine("App.config");
                writer.WriteLine(File.ReadAllText(appConfigPath));
                writer.WriteLine(new string('*', 65)); // separator
                writer.WriteLine();
            }

            // Collect all .cs files
            string[] csFiles = Directory.GetFiles(searchFolderPath, "*.cs", SearchOption.TopDirectoryOnly);
            foreach (string csFile in csFiles)
            {
                string fileName = Path.GetFileName(csFile);
                writer.WriteLine(fileName);
                writer.WriteLine(File.ReadAllText(csFile));
                writer.WriteLine(new string('*', 65)); // separator
                writer.WriteLine();
            }
        }

        Console.WriteLine("Contents collected into: " + outputFilePath);
    }
}