using System.Threading.Tasks;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Task = System.Threading.Tasks.Task;

namespace BlazorApp3.Data
{
    public class InputService
    {
        public Task<string> GetInputExampleAsync()
        {
            return Task.FromResult(File.ReadAllText("wwwroot\\Settings\\InputExample.txt"));
        }

        public Task<string> GetInputExampleAsync(string fileName)
        {
            string path = $"wwwroot\\Upload\\{fileName}";

            if (new FileInfo(path).Extension.Equals(".txt"))
            {
                return Task.FromResult(File.ReadAllText(path));
            }
            else
            {
                using (FileStream fileStreamPath = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic);
                    return Task.FromResult(document.GetText().Remove(0, 61));
                }
            }           
        }

        public void DeleteFile(string fileName)
        {
            if (File.Exists($"wwwroot\\Upload\\{fileName}"))
            { 
                File.Delete($"wwwroot\\Upload\\{fileName}");
            }
        }
    }
}
