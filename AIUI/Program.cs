using System.IO;
using Microsoft.Office.Interop.Word;

public void ConvertDocToDocx(string path)
{
    Application word = new Application();

    if (path.ToLower().EndsWith(".doc"))
    {
        var sourceFile = new FileInfo(path);
        var document = word.Documents.Open(sourceFile.FullName);

        // Define the new file name with the .docx extension
        string newFileName = Path.ChangeExtension(sourceFile.FullName, ".docx");

        // Save the document in the .docx format
        document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument);

        // Close and release resources
        document.Close();
        word.Quit();
        
        // Delete the original .doc file if needed
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
