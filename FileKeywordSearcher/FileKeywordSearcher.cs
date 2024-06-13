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
            string strLineMapping = "";
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
                if (CheckFileForKeyword(file, ref strLineMapping))
                {
                    FileItem fileItem = new FileItem(file, true, strLineMapping);
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

        private bool CheckFileForKeyword(string filePath, ref string strLineMapping)
        {
            bool bHasKeyWord = false;
            List<int> keywordLines = new List<int>();

            try
            {
                // Read all lines from the file
                string[] lines = File.ReadAllLines(filePath);

                // Loop through each line in the file
                for (int i = 0; i < lines.Length; i++)
                {
                    // Check if the current line contains the keyword (case insensitive)
                    if (lines[i].IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // If the keyword is found in the line, add the line number to the list
                        keywordLines.Add(i + 1); // Add 1 because line numbers start from 1
                        bHasKeyWord = true;
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions such as file not found, access denied, etc.
                MessageBox.Show($"Error reading file {filePath}: {ex.Message}");
            }

            // Check if any keyword was found in the file
            if (bHasKeyWord)
            {
                // If keywords were found, convert the list of line numbers to a string
                strLineMapping = string.Join(", ", keywordLines);
            }
            else
            {
                // If no keyword was found, set strLineMapping to an empty string
                strLineMapping = "";
            }

            return bHasKeyWord;
        }


    }
}
