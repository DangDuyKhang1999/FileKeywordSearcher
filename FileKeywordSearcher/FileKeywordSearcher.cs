using System.Text;
using static FileKeywordSearcher.Form1;
using ClosedXML.Excel;

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
            HasKeyWord(m_strBrowser);
        }

        public List<FileItem> GetFileItems()
        {
            return m_fileItems;
        }

        public bool HasKeyWord(string directoryPath)
        {
            bool iResult = false;

            // Check all files in the current directory
            foreach (var file in Directory.GetFiles(directoryPath, "*"))
            {
                string strLineMapping = "";
                FileExtension fileExtension = GetFileExtension(file);
                if (fileExtension == FileExtension.Normal)
                {
                    if (CheckFileForKeyword(file, ref strLineMapping))
                    {
                        FileItem fileItem = new FileItem(file, strLineMapping, fileExtension);
                        m_fileItems.Add(fileItem);
                        iResult = true; // If at least one file is found, set result to true
                    }
                }
                else if (fileExtension == FileExtension.CSV)
                {
                    if (CheckCSVForKeyword(file, ref strLineMapping))
                    {
                        FileItem fileItem = new FileItem(file, strLineMapping, fileExtension);
                        m_fileItems.Add(fileItem);
                        iResult = true; // If at least one file is found, set result to true
                    }
                }
                else if (fileExtension == FileExtension.Excel)
                {
                    if (CheckExcelForKeyword(file, ref strLineMapping))
                    {
                        FileItem fileItem = new FileItem(file, strLineMapping, fileExtension);
                        m_fileItems.Add(fileItem);
                        iResult = true; // If at least one file is found, set result to true
                    }
                }
            }

            // Recursively check all subdirectories
            foreach (var subDirectory in Directory.GetDirectories(directoryPath))
            {
                iResult |= HasKeyWord(subDirectory);
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

        private bool CheckCSVForKeyword(string filePath, ref string strCellMapping)
        {
            bool bHasKeyWord = false;
            List<string> keywordCells = new List<string>();

            try
            {
                // Read all lines from the CSV file
                string[] lines = File.ReadAllLines(filePath);

                // Loop through each line in the file
                for (int i = 0; i < lines.Length; i++)
                {
                    // Split the line into cells (assuming comma as delimiter)
                    string[] cells = lines[i].Split(',');

                    // Loop through each cell in the line
                    for (int j = 0; j < cells.Length; j++)
                    {
                        // Check if the current cell contains the keyword (case insensitive)
                        if (cells[j].IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // If the keyword is found in the cell, add the cell position to the list
                            string cellPosition = $"{GetExcelColumnName(j + 1)}{i + 1}";
                            keywordCells.Add(cellPosition);
                            bHasKeyWord = true;
                        }
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
                // If keywords were found, convert the list of cell positions to a string
                strCellMapping = string.Join(", ", keywordCells);
            }
            else
            {
                // If no keyword was found, set strCellMapping to an empty string
                strCellMapping = "";
            }

            return bHasKeyWord;
        }

        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = String.Empty;
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }
        private FileExtension GetFileExtension(string fileName)
        {
            string extension = Path.GetExtension(fileName).ToLowerInvariant();
            return extension switch
            {
                ".xls" => FileExtension.Excel,
                ".xlsx" => FileExtension.Excel,
                ".csv" => FileExtension.CSV,
                _ => FileExtension.Normal
            };
        }

        public bool CheckExcelForKeyword(string filePath, ref string strCellMapping)
        {
            bool bHasKeyWord = false;
            List<string> keywordCells = new List<string>();

            try
            {
                // Open the Excel workbook
                using (var workbook = new XLWorkbook(filePath))
                {
                    // Loop through each worksheet in the workbook
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        // Loop through each row in the worksheet
                        foreach (var row in worksheet.RowsUsed())
                        {
                            // Loop through each cell in the row
                            foreach (var cell in row.CellsUsed())
                            {
                                // Check if the current cell contains the keyword (case insensitive)
                                if (cell.GetString().IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // If the keyword is found in the cell, add the cell address to the list
                                    keywordCells.Add(cell.Address.ToString());
                                    bHasKeyWord = true;
                                }
                            }
                        }
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
                // If keywords were found, convert the list of cell positions to a string
                strCellMapping = string.Join(", ", keywordCells);
            }
            else
            {
                // If no keyword was found, set strCellMapping to an empty string
                strCellMapping = "";
            }

            return bHasKeyWord;
        }
    }
}
