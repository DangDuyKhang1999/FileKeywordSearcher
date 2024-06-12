using System.Text;
using static FileKeywordSearcher.Form1;

namespace FileKeywordSearcher
{
    internal class FileKeywordSearcher
    {
        public string m_strBrowser { get; set; }
        public string m_strKeyWord { get; set; }

        private List<FileItem> m_fileItems = new List<FileItem>();

        public FileKeywordSearcher(string strBrowser, string strKeyWord)
        {
            m_strBrowser = strBrowser;
            m_strKeyWord = strKeyWord;
            HasCodeFiles(m_strBrowser);
        }

        public List<FileItem> GetFileItems()
        {
            return m_fileItems;
        }

        public bool HasCodeFiles(string directoryPath)
        {
            bool iResult = false;

            // Check source files in the current directory
            foreach (var file in Directory.GetFiles(directoryPath, "*.h") //Header C++
                                             .Concat(Directory.GetFiles(directoryPath, "*.cpp"))   // C++
                                             .Concat(Directory.GetFiles(directoryPath, "*.c"))     // C
                                             .Concat(Directory.GetFiles(directoryPath, "*.cs"))    //C#
                                             .Concat(Directory.GetFiles(directoryPath, "*.java"))  //Java
                                             .Concat(Directory.GetFiles(directoryPath, "*.py"))    // Python
                                             .Concat(Directory.GetFiles(directoryPath, "*.rb"))    // Ruby
                                             .Concat(Directory.GetFiles(directoryPath, "*.php"))   // PHP
                                             .Concat(Directory.GetFiles(directoryPath, "*.swift")) // Swift
                                             .Concat(Directory.GetFiles(directoryPath, "*.go"))    // Go
                                             .Concat(Directory.GetFiles(directoryPath, "*.ts"))    // TypeScript
                                             .Concat(Directory.GetFiles(directoryPath, "*.kt"))    // Kotlin
                                             .Concat(Directory.GetFiles(directoryPath, "*.scala")) // Scala
                                             .Concat(Directory.GetFiles(directoryPath, "*.pl"))    // Perl
                                             .Concat(Directory.GetFiles(directoryPath, "*.lua"))   // Lua
                                             .Concat(Directory.GetFiles(directoryPath, "*.dart"))  // Dart (Flutter)
                                             .Concat(Directory.GetFiles(directoryPath, "*.js"))    // JavaScript (React Native)
                                             .Concat(Directory.GetFiles(directoryPath, "*.jsx"))   // JSX (React Native)
                                             .Concat(Directory.GetFiles(directoryPath, "*.m"))     //MATLAB
                                             .Concat(Directory.GetFiles(directoryPath, "*.csv"))   //CSV
                                             .Concat(Directory.GetFiles(directoryPath, "*.txt"))   //text
                                             )
            {
                if (CheckFileForKeyword(file))
                { 
                    FileItem fileItem = new FileItem(file, true);
                    m_fileItems.Add(fileItem);
                }
            }

            // Recursively check all subdirectories
            foreach (var subDirectory in Directory.GetDirectories(directoryPath))
            {
                iResult |= HasCodeFiles(subDirectory);
            }

            return iResult;
        }

        private bool CheckFileForKeyword(string filePath)
        {
            bool bHasKeyWord = false;
            try
            {
                string content = File.ReadAllText(filePath);
                if (content.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    bHasKeyWord = true;
                }
                else
                {
                    bHasKeyWord = false;
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions such as file not found, access denied, etc.
                MessageBox.Show($"Error reading file {filePath}: {ex.Message}");
            }
            return bHasKeyWord;
        }

    }
}
