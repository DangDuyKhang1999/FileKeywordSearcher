using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileKeywordSearcher
{
    /// <summary>
    /// File extension of each file found in the folder (FileItem.m_fileExtension)
    /// </summary>
    public enum eFileExtension
    {
        Normal,
        CSV,
        Excel,
        Excel_Old,
        Word,
        Word_Old,
        Word_RTF,
        PowerPoint,
        PowerPoint_old,
        PDF,
        IgnoredExtension,
    }

    /// <summary>
    /// Enum to create items for the target extension selection listbox (control: labelWithCheckBoxList)
    /// </summary>
    public enum eTargetExtension
    {
        PlainText,  // Plain text file format (e.g., .txt, .asc, .log, .csv)
        CSV,        // Comma-separated values file format (e.g., .csv)
        Excel,      // Microsoft Excel file format (e.g., .xls, .xlsx)
        Word,       // Microsoft Word file format (e.g., .doc, .docx)
        PowerPoint, // Microsoft PowerPoint file format (e.g., .ppt, .pptx)
        PDF,        // Portable Document Format (e.g., .pdf)
    }

}
