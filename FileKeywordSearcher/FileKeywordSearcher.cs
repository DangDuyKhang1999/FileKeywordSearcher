using System.Text;
using static FileKeywordSearcher.Form1;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Path = System.IO.Path;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace FileKeywordSearcher
{
    internal class FileKeywordSearcher
    {
        public string m_strBrowser { get; set; }
        public string m_strKeyWord { get; set; }
        public int m_iFileCount { get; set; }
        public int m_iTotalFileCount { get; set; }

        private List<FileItem> m_fileItems = new List<FileItem>();

        public event EventHandler<int> ProgressChanged;

        public FileKeywordSearcher(string strBrowser, string strKeyWord)
        {
            m_iFileCount = 0;
            m_iTotalFileCount = 0;
            m_strBrowser = strBrowser;
            m_strKeyWord = strKeyWord;
        }

        public List<FileItem> GetFileItems()
        {
            return m_fileItems;
        }
        protected virtual void OnProgressChanged(int percent)
        {
            ProgressChanged?.Invoke(this, percent); // Trigger the ProgressChanged event
        }
        public bool HasKeyWord(string directoryPath)
        {
            bool iResult = false;

            // Check all files in the current directory
            foreach (var file in Directory.GetFiles(directoryPath, "*"))
            {
                string strLineMapping = "";
                FileExtension fileExtension = GetFileExtension(file);
                bool keywordFound = false;
                bool bHasMultiKeyWord = false;

                switch (fileExtension)
                {
                    case FileExtension.Normal:
                        keywordFound = CheckFileForKeyword(file, ref strLineMapping, ref bHasMultiKeyWord);
                        break;

                    case FileExtension.CSV:
                        keywordFound = CheckCSVForKeyword(file, ref strLineMapping, ref bHasMultiKeyWord);
                        break;

                    case FileExtension.Excel:
                        keywordFound = CheckExcelForKeywordAndShapes(file, ref strLineMapping);
                        break;

                    case FileExtension.PDF:
                        keywordFound = CheckPDFForKeyword(file, ref strLineMapping, ref bHasMultiKeyWord);
                        break;
                }

                if (keywordFound)
                {
                    FileItem fileItem = new FileItem(file, strLineMapping, fileExtension, bHasMultiKeyWord);
                    m_fileItems.Add(fileItem);
                    iResult = true; // If at least one file is found, set result to true
                }
                m_iFileCount++;
                int percentComplete = (int)((double)m_iFileCount / m_iTotalFileCount * 100);
                OnProgressChanged(percentComplete);
            }

            // Recursively check all subdirectories
            foreach (var subDirectory in Directory.GetDirectories(directoryPath))
            {
                iResult |= HasKeyWord(subDirectory);
            }

            return iResult;
        }
        public int CountFiles(string directoryPath)
        {
            int totalCount = 0;

            // Count all files in the current directory
            totalCount += Directory.GetFiles(directoryPath).Length;

            // Recursively count all files in subdirectories
            foreach (var subDirectory in Directory.GetDirectories(directoryPath))
            {
                totalCount += CountFiles(subDirectory);
            }
            return totalCount;
        }
        public bool getTotalFiles()
        {
            try
            {
                m_iTotalFileCount = CountFiles(m_strBrowser);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show($"Access to directory '{m_strBrowser}' is denied: {ex.Message}. Please ensure you have the necessary permissions to access this folder.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Exception occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (m_iTotalFileCount == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool CheckFileForKeyword(string filePath, ref string strLineMapping, ref bool bHasMultiKeyWord)
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
            if (keywordLines.Count > 1)
            {
                bHasMultiKeyWord = true;
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

        private bool CheckCSVForKeyword(string filePath, ref string strCellMapping, ref bool bHasMultiKeyWord)
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
            if (keywordCells.Count > 1)
            {
                bHasMultiKeyWord = true;
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

        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (document != null)
            {
                SharedStringTablePart? sstPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                if (sstPart != null)
                {
                    string value = cell.InnerText;

                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        int index;
                        if (Int32.TryParse(value, out index) && index < sstPart.SharedStringTable.ChildElements.Count)
                        {
                            return sstPart.SharedStringTable.ChildElements[index].InnerText;
                        }
                        else
                        {
                            return value;
                        }
                    }
                    else
                    {
                        return value;
                    }
                }

            }
            return "";
        }

        public bool CheckExcelForKeywordAndShapes(string filePath, ref string strMapping)
        {
            bool bHasKeyWord = false;
            Dictionary<string, List<string>> keywordCells = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> sheetShapes = new Dictionary<string, List<string>>();

            try
            {
                // Open the Excel document
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
                {
                    if (document == null || document.WorkbookPart == null)
                    {
                        MessageBox.Show($"Unable to open or read the Excel document at {filePath}");
                        return false; // Exit early if document or workbook part is null
                    }

                    WorkbookPart workbookPart = document.WorkbookPart;
                    if (workbookPart == null || workbookPart.Workbook == null || workbookPart.Workbook.Sheets == null)
                    {
                        MessageBox.Show($"Unable to open or read the Excel document at {filePath}");
                        return false; // Exit early if document or workbook part is null
                    }

                    foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                    {
                        if (sheet == null || workbookPart == null)
                        {
                            continue; // Skip null sheets or workbook parts
                        }
                        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                        string sheetName = sheet.Name;

                        // Initialize the list of cells containing keyword in the current worksheet
                        if (!keywordCells.ContainsKey(sheetName))
                        {
                            keywordCells[sheetName] = new List<string>();
                        }

                        // Initialize the list of shapes in the current worksheet
                        if (!sheetShapes.ContainsKey(sheetName))
                        {
                            sheetShapes[sheetName] = new List<string>();
                        }

                        // Check keyword in cells
                        if (worksheetPart == null || worksheetPart.Worksheet == null)
                        {
                            continue; // Skip null worksheet parts or worksheets
                        }
                        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            foreach (Row row in sheetData.Elements<Row>())
                            {
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    if (cell == null || cell.CellReference == null)
                                    {
                                        continue; // Skip null cell references
                                    }

                                    // Get cell value, checking for null
                                    string cellValue = GetCellValue(document, cell);
                                    if (cellValue == null)
                                    {
                                        continue; // Skip null cell values
                                    }

                                    // Check if the cell contains the keyword (case insensitive)
                                    if (cellValue.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        string cellAddress = cell.CellReference.ToString();
                                        if (cellAddress != null)
                                        {
                                            if (!keywordCells[sheetName].Contains(cellAddress))
                                            {
                                                keywordCells[sheetName].Add(cellAddress);
                                                bHasKeyWord = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Check shapes in worksheet
                        if (worksheetPart.DrawingsPart != null)
                        {
                            var drawingsPart = worksheetPart.DrawingsPart;
                            var shapeElements = drawingsPart.WorksheetDrawing.Elements<TwoCellAnchor>();

                            foreach (var element in shapeElements)
                            {
                                if (element == null || element.FromMarker == null || element.FromMarker.RowId == null || element.FromMarker.ColumnId == null)
                                {
                                    continue; // Skip null elements or markers
                                }

                                // Get the text content of the shape
                                var shapeText = element.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text).Aggregate(string.Empty, (current, text) => current + text);

                                // Get the start position of the shape
                                int fromRow = int.Parse(element.FromMarker.RowId.Text); // Row index (0-based)
                                int fromColumn = int.Parse(element.FromMarker.ColumnId.Text); // Column index (0-based)
                                string shapePosition = $"{GetExcelColumnName(fromColumn + 1)}{fromRow + 1}"; // Convert to 1-based

                                // Check if the shape text contains the keyword (case insensitive)
                                if (shapeText.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    if (sheetShapes.ContainsKey(sheetName) && !sheetShapes[sheetName].Contains(shapePosition))
                                    {
                                        sheetShapes[sheetName].Add(shapePosition);
                                        bHasKeyWord = true;
                                    }
                                }
                            }
                        }
                    }
                }

                // Build strMapping combining keyword cells and shapes for each sheet
                List<string> resultMappings = new List<string>();
                foreach (var kvp in keywordCells)
                {
                    string sheetName = kvp.Key;
                    List<string> cellsInSheet = kvp.Value;
                    List<string> shapesInSheet = sheetShapes.ContainsKey(sheetName) ? sheetShapes[sheetName] : new List<string>();

                    // Combine cells and shapes for the sheet into a single string
                    string sheetMapping = $"\"{sheetName}\":: Cells({string.Join(", ", cellsInSheet)}), Shapes({string.Join(", ", shapesInSheet)})";
                    resultMappings.Add(sheetMapping);
                }

                // Update strMapping with the combined mappings
                strMapping = string.Join("; ", resultMappings);
            }
            catch (Exception ex)
            {
                // Handle exceptions such as file not found, access denied, etc.
                MessageBox.Show($"Error reading file {filePath}: {ex.Message}");
                return false;
            }

            // Return whether the keyword was found or not
            return bHasKeyWord;
        }

        public bool CheckPDFForKeyword(string filePath, ref string strKeywordMapping, ref bool bHasMultiKeyWord)
        {
            bool bHasKeyword = false;
            HashSet<int> pagesWithKeyword = new HashSet<int>();

            try
            {
                // Open the PDF file
                using (PdfReader reader = new PdfReader(filePath))
                {
                    // Iterate through each page in the PDF file
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        string text = PdfTextExtractor.GetTextFromPage(reader, i);

                        // Process the text to remove unwanted spaces
                        text = text.Replace(" ", "");

                        // Search for the keyword in the text (case insensitive)
                        if (text.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Save the page number if the keyword is found
                            pagesWithKeyword.Add(i);
                            bHasKeyword = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions such as file not found, access denied, etc.
                MessageBox.Show($"Error reading file {filePath}: {ex.Message}");
            }

            // Build strKeywordMapping from the set of pagesWithKeyword
            if (bHasKeyword)
            {
                if (pagesWithKeyword.Count > 1)
                {
                    bHasMultiKeyWord = true;
                }
                string newKeywordMapping = string.Join(", ", pagesWithKeyword);

                // Update strKeywordMapping only if the keyword is found
                if (!string.IsNullOrEmpty(strKeywordMapping))
                {
                    strKeywordMapping += "; " + newKeywordMapping;
                }
                else
                {
                    strKeywordMapping = newKeywordMapping; // Set to newKeywordMapping if strKeywordMapping is empty
                }
            }

            // Return whether the keyword was found or not
            return bHasKeyword;
        }

        private FileExtension GetFileExtension(string fileName)
        {
            string extension = Path.GetExtension(fileName).ToLowerInvariant();
            return extension switch
            {
                ".xls" => FileExtension.Excel,
                ".xlsx" => FileExtension.Excel,
                ".csv" => FileExtension.CSV,
                ".pdf" => FileExtension.PDF,
                _ => FileExtension.Normal
            };
        }
    }
}
