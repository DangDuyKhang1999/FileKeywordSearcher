using System.Text;
using static FileKeywordSearcher.Form1;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Path = System.IO.Path;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HWPF;
using NPOI.HWPF.UserModel;


namespace FileKeywordSearcher
{
    internal class FileKeywordSearcher
    {
        public string m_strBrowser { get; set; }
        public string m_strKeyWord { get; set; }
        public int m_iFileCount { get; set; }
        public int m_iTotalFileCount { get; set; }
        public List<FileItem> m_fileItems { get; set; } = new List<FileItem>();


        public event EventHandler<(int percent, int iFileCount, int iTotalFileCount, int iFileHasKeyWord)> ProgressChanged;

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
        protected virtual void OnProgressChanged(int percent, string filePath)
        {
            ProgressChanged?.Invoke(this, (percent, m_iFileCount, m_iTotalFileCount, m_fileItems.Count)); // Trigger the ProgressChanged event with percent and filePath
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

                    case FileExtension.Excel_Old:
                        keywordFound = CheckOldExcelForKeywordAndShapes(file, ref strLineMapping);
                        break;

                    case FileExtension.PDF:
                        keywordFound = CheckPDFForKeyword(file, ref strLineMapping, ref bHasMultiKeyWord);
                        break;

                    case FileExtension.Word:
                        keywordFound = CheckWordForKeywordAndShapes(file);
                        break;
                    
                    case FileExtension.Word_Old:
                        keywordFound = CheckOldWordForKeywordAndShapes(file);
                        break;

                    case FileExtension.Word_RTF:
                        keywordFound = CheckFileForKeyword(file, ref strLineMapping, ref bHasMultiKeyWord);
                        break;

                    case FileExtension.PowerPoint:
                        keywordFound = CheckPowerPointForKeywordAndShapes(file);
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
                OnProgressChanged(percentComplete, directoryPath);
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
        //Excel ---------->
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
                    if (workbookPart == null || workbookPart.Workbook.Sheets == null)
                    {
                        MessageBox.Show($"Unable to open or read the Excel document at {filePath}");
                        return false; // Exit early if document or workbook part is null
                    }
                    foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                    {
                        if (sheet == null || workbookPart == null)
                        {
                            return false; // Exit early if document or workbook part is null
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
                            return false; // Exit early if document or workbook part is null
                        }
                        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            foreach (Row row in sheetData.Elements<Row>())
                            {
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    if (cell != null && cell.CellReference != null)
                                    {
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
                        }

                        // Check shapes in worksheet
                        if (worksheetPart.DrawingsPart != null)
                        {
                            var drawingsPart = worksheetPart.DrawingsPart;
                            var shapeElements = drawingsPart.WorksheetDrawing.Elements<TwoCellAnchor>();

                            foreach (var element in shapeElements)
                            {
                                // Get the text content of the shape
                                var shapeText = element.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text).Aggregate(string.Empty, (current, text) => current + text);

                                // Get the start position of the shape
                                var fromMarker = element.FromMarker;
                                if (fromMarker != null && fromMarker.RowId != null && fromMarker.ColumnId != null)
                                {
                                    int fromRow = int.Parse(fromMarker.RowId.Text); // Row index (0-based)
                                    int fromColumn = int.Parse(fromMarker.ColumnId.Text); // Column index (0-based)
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
                }

                // Build strMapping combining keyword cells and shapes for each sheet
                List<string> resultMappings = new List<string>();
                foreach (var kvp in keywordCells)
                {
                    string sheetName = kvp.Key;
                    List<string> cellsInSheet = kvp.Value;
                    List<string> shapesInSheet = sheetShapes.TryGetValue(sheetName, out var shapes) ? shapes : new List<string>();
                    bool bHasKeyWordInSheet = false;
                    string sheetMapping = $"{sheetName}:: ";

                    if (cellsInSheet.Count > 0)
                    {
                        sheetMapping += (cellsInSheet.Count > 1) ? $"Cells({string.Join(", ", cellsInSheet)})" : $"Cell({cellsInSheet[0]})";
                        bHasKeyWordInSheet = true;
                    }

                    if (shapesInSheet.Count > 0)
                    {
                        if (sheetMapping.Length > sheetName.Length + 3)
                        {
                            sheetMapping += ", ";
                        }
                        sheetMapping += (shapesInSheet.Count > 1) ? $"Shapes({string.Join(", ", shapesInSheet)})" : $"Shape({shapesInSheet[0]})";
                        bHasKeyWordInSheet = true;
                    }
                    if (bHasKeyWordInSheet) { resultMappings.Add(sheetMapping); }
                }

                // Update strMapping with the combined mappings
                strMapping = string.Join("; ", resultMappings);
            }
            catch (Exception ex)
            {
                // Handle exceptions such as file not found, access denied, etc.
                MessageBox.Show($"Error reading file {filePath}: {ex.Message}");
            }

            // Return whether the keyword was found or not
            return bHasKeyWord;
        }

        public bool CheckOldExcelForKeywordAndShapes(string filePath, ref string strMapping)
        {
            bool bHasKeyWord = false;
            Dictionary<string, List<string>> keywordCells = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> sheetShapes = new Dictionary<string, List<string>>();

            try
            {
                HSSFWorkbook workbook;
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = new HSSFWorkbook(fileStream);
                }

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    ISheet sheet = workbook.GetSheetAt(i);
                    string sheetName = sheet.SheetName;

                    // Initialize list of cells containing the keyword in the current worksheet
                    if (!keywordCells.ContainsKey(sheetName))
                    {
                        keywordCells[sheetName] = new List<string>();
                    }

                    // Initialize list of shapes in the current worksheet
                    if (!sheetShapes.ContainsKey(sheetName))
                    {
                        sheetShapes[sheetName] = new List<string>();
                    }

                    // Check for keyword in cells
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null) continue;

                        for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                        {
                            ICell cell = row.GetCell(cellIndex);
                            if (cell == null) continue;

                            string cellValue = cell.ToString();
                            if (cellValue.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                string cellAddress = $"{GetExcelColumnName(cellIndex + 1)}{rowIndex + 1}";
                                if (!keywordCells[sheetName].Contains(cellAddress))
                                {
                                    keywordCells[sheetName].Add(cellAddress);
                                    bHasKeyWord = true;
                                }
                            }
                        }
                    }

                    // Check for shapes in the worksheet
                    HSSFPatriarch drawingPatriarch = sheet.DrawingPatriarch as HSSFPatriarch;
                    if (drawingPatriarch != null)
                    {
                        foreach (HSSFShape shape in drawingPatriarch.Children)
                        {
                            if (shape is HSSFSimpleShape simpleShape && simpleShape.String != null)
                            {
                                string shapeText = simpleShape.String.String;
                                if (shapeText.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    HSSFClientAnchor anchor = shape.Anchor as HSSFClientAnchor;
                                    if (anchor != null)
                                    {
                                        string shapePosition = $"{GetExcelColumnName(anchor.Col1 + 1)}{anchor.Row1 + 1}";
                                        if (!sheetShapes[sheetName].Contains(shapePosition))
                                        {
                                            sheetShapes[sheetName].Add(shapePosition);
                                            bHasKeyWord = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // Build strMapping combining cells containing the keyword and shapes for each sheet
                List<string> resultMappings = new List<string>();
                foreach (var kvp in keywordCells)
                {
                    string sheetName = kvp.Key;
                    List<string> cellsInSheet = kvp.Value;
                    List<string> shapesInSheet = sheetShapes.TryGetValue(sheetName, out var shapes) ? shapes : new List<string>();
                    bool bHasKeyWordInSheet = false;
                    string sheetMapping = $"{sheetName}:: ";

                    if (cellsInSheet.Count > 0)
                    {
                        sheetMapping += (cellsInSheet.Count > 1) ? $"Cells({string.Join(", ", cellsInSheet)})" : $"Cell({cellsInSheet[0]})";
                        bHasKeyWordInSheet = true;
                    }

                    if (shapesInSheet.Count > 0)
                    {
                        if (sheetMapping.Length > sheetName.Length + 3)
                        {
                            sheetMapping += ", ";
                        }
                        sheetMapping += (shapesInSheet.Count > 1) ? $"Shapes({string.Join(", ", shapesInSheet)})" : $"Shape({shapesInSheet[0]})";
                        bHasKeyWordInSheet = true;
                    }

                    if (bHasKeyWordInSheet)
                    {
                        resultMappings.Add(sheetMapping);
                    }
                }

                // Update strMapping with combined results
                strMapping = string.Join("; ", resultMappings);
            }
            catch (Exception ex)
            {
                // Handle exceptions such as file not found, access denied, etc.
                MessageBox.Show($"Error reading file {filePath}: {ex.Message}");
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
        //Excel <----------
        //Word ---------->
        public bool CheckWordForKeywordAndShapes(string filePath)
        {
            bool bHasKeyWord = false;
            Dictionary<string, List<string>> keywordTexts = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> shapePositions = new Dictionary<string, List<string>>();

            try
            {
                // Open the Word document
                using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, false))
                {
                    if (document == null || document.MainDocumentPart == null)
                    {
                        MessageBox.Show($"Unable to open or read the Word document at {filePath}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false; // Exit early if document or main part is null
                    }

                    MainDocumentPart mainPart = document.MainDocumentPart;

                    // Iterate through paragraphs to check for keyword
                    foreach (var paragraph in mainPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                    {
                        string paragraphText = paragraph.InnerText;

                        // Check if paragraph contains the keyword (case insensitive)
                        if (paragraphText.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            bHasKeyWord = true;
                            break; // Exit loop early if keyword is found
                        }
                    }

                    // Iterate through drawings to check for shapes
                    foreach (var drawing in mainPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
                    {
                        var inline = drawing.Inline;

                        // Check if the drawing contains text and if that text contains the keyword
                        if (inline != null)
                        {
                            var drawingText = inline.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text).Aggregate(string.Empty, (current, text) => current + text);
                            if (drawingText.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                bHasKeyWord = true;
                                break; // Exit loop early if keyword is found
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions such as file not found, access denied, etc.
                MessageBox.Show($"Error reading file {filePath}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // Return whether the keyword was found or not
            return bHasKeyWord;
        }
        public bool CheckOldWordForKeywordAndShapes(string filePath)
        {
            bool bHasKeyWord = false;
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    HWPFDocument doc = new HWPFDocument(fs);

                    // Check document text for keyword
                    string documentText = doc.GetDocumentText();
                    if (documentText.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        bHasKeyWord = true;
                    }
                }

                return bHasKeyWord;
            }
            catch (Exception ex)
            {
                // Handle exceptions such as file not found, access denied, etc.
                MessageBox.Show($"Error reading file {filePath}: {ex.Message}");
                return false; // Return false if an error occurs
            }
        }
        //Word <----------

        //PowerPoint ---------->
        public bool CheckPowerPointForKeywordAndShapes(string filePath)
        {
            bool hasKeyword = false;

            try
            {
                // Open the PowerPoint presentation
                using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, false))
                {
                    if (presentationDocument == null || presentationDocument.PresentationPart == null)
                    {
                        MessageBox.Show($"Unable to open or read the PowerPoint presentation at {filePath}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false; // Exit early if document or main part is null
                    }

                    PresentationPart presentationPart = presentationDocument.PresentationPart;
                    Presentation presentation = presentationPart.Presentation;

                    // Iterate through slides
                    foreach (SlideId slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);

                        // Check for keyword in slide text
                        foreach (DocumentFormat.OpenXml.Presentation.Shape shape in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>())
                        {
                            if (shape.TextBody != null)
                            {
                                string shapeText = shape.TextBody.InnerText;

                                // Check if the shape contains the keyword (case insensitive)
                                if (shapeText.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    hasKeyword = true;
                                    break; // Exit loop early if keyword is found
                                }
                            }
                        }

                        if (hasKeyword)
                        {
                            break; // Exit outer loop if keyword is found
                        }

                        // Check for shapes in slide drawings
                        foreach (var graphicFrame in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>())
                        {
                            var drawingTexts = graphicFrame.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text);
                            string drawingText = string.Join("", drawingTexts);

                            if (drawingText.IndexOf(m_strKeyWord, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                hasKeyword = true;
                                break; // Exit loop early if keyword is found
                            }
                        }

                        if (hasKeyword)
                        {
                            break; // Exit outer loop if keyword is found
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions such as file not found, access denied, etc.
                MessageBox.Show($"Error reading file {filePath}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // Return whether the keyword was found or not
            return hasKeyword;
        }
        //PowerPoint <----------
        private FileExtension GetFileExtension(string fileName)
        {
            string extension = Path.GetExtension(fileName).ToLowerInvariant();
            return extension switch
            {
                ".csv" => FileExtension.CSV, // Comma-Separated Values
                ".xls" => FileExtension.Excel_Old, // Microsoft Excel Spreadsheet (Legacy)
                ".xlsx" => FileExtension.Excel, // Microsoft Excel Spreadsheet
                ".doc" => FileExtension.Word_Old, // Microsoft Word document  (Legacy)
                ".docx" => FileExtension.Word, // Microsoft Word document
                ".docm" => FileExtension.Word, // Microsoft Word document with macros
                ".rtf" => FileExtension.Word_RTF, // Microsoft Word document in Rich Text Format (RTF)
                ".ppt" => FileExtension.PowerPoint_old, // Microsoft PowerPoint presentation (Legacy)
                ".pptm" => FileExtension.PowerPoint, // Microsoft PowerPoint presentation with macros
                ".pptx" => FileExtension.PowerPoint, // Microsoft PowerPoint presentation (Open XML format)
                ".pdf" => FileExtension.PDF, // Portable Document Format
                // Ignored Extension ----------->
                ".exe" => FileExtension.IgnoredExtension, // Executable File
                ".nupkg" => FileExtension.IgnoredExtension, // NuGet Package
                ".dll" => FileExtension.IgnoredExtension, // Dynamic Link Library
                ".bin" => FileExtension.IgnoredExtension, // Binary File
                ".img" => FileExtension.IgnoredExtension, // Disk Image File
                ".iso" => FileExtension.IgnoredExtension, // Optical Disc Image
                ".jpg" => FileExtension.IgnoredExtension, // JPEG Image
                ".jpeg" => FileExtension.IgnoredExtension, // JPEG Image 
                ".png" => FileExtension.IgnoredExtension, // Portable Network Graphics
                ".gif" => FileExtension.IgnoredExtension, // Graphics Interchange Format
                ".bmp" => FileExtension.IgnoredExtension, // Bitmap Image
                ".tiff" => FileExtension.IgnoredExtension, // Tagged Image File Format
                ".mp3" => FileExtension.IgnoredExtension, // MPEG Audio Layer III
                ".wav" => FileExtension.IgnoredExtension, // Waveform Audio File Format
                ".flac" => FileExtension.IgnoredExtension, // Free Lossless Audio Codec
                ".aac" => FileExtension.IgnoredExtension, // Advanced Audio Codec
                ".ogg" => FileExtension.IgnoredExtension, // Ogg Vorbis
                ".mp4" => FileExtension.IgnoredExtension, // MPEG-4 Video
                ".mkv" => FileExtension.IgnoredExtension, // Matroska Video
                ".avi" => FileExtension.IgnoredExtension, // Audio Video Interleave
                ".mov" => FileExtension.IgnoredExtension, // QuickTime Movie
                ".wmv" => FileExtension.IgnoredExtension, // Windows Media Video
                ".zip" => FileExtension.IgnoredExtension, // ZIP Archive
                ".rar" => FileExtension.IgnoredExtension, // RAR Archive
                ".7z" => FileExtension.IgnoredExtension, // 7-Zip Archive
                ".tar.gz" => FileExtension.IgnoredExtension, // Compressed Archive File
                ".db" => FileExtension.IgnoredExtension, // Database File
                ".mdb" => FileExtension.IgnoredExtension, // Microsoft Access Database
                ".sqlite" => FileExtension.IgnoredExtension, // SQLite Database
                ".psd" => FileExtension.IgnoredExtension, // Adobe Photoshop Document
                ".ai" => FileExtension.IgnoredExtension, // Adobe Illustrator Document
                ".dwg" => FileExtension.IgnoredExtension, // AutoCAD Drawing
                ".sys" => FileExtension.IgnoredExtension, // System File
                ".dat" => FileExtension.IgnoredExtension, // Data File
                ".wma" => FileExtension.IgnoredExtension, // Windows Media Audio
                ".ps" => FileExtension.IgnoredExtension, // PostScript File
                ".key" => FileExtension.IgnoredExtension, // Keynote Presentation
                ".odt" => FileExtension.IgnoredExtension, // OpenDocument Text Document
                ".ods" => FileExtension.IgnoredExtension, // OpenDocument Spreadsheet
                ".odp" => FileExtension.IgnoredExtension, // OpenDocument Presentation
                ".dwf" => FileExtension.IgnoredExtension, // Design Web Format
                ".jar" => FileExtension.IgnoredExtension, // Java Archive
                // ----------->
                _ => FileExtension.Normal
            };
        }
    }
}
