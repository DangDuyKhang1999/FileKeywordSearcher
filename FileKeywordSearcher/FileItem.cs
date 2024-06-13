using System;
using static FileKeywordSearcher.Form1;

namespace FileKeywordSearcher
{
    internal class FileItem
    {
        public string m_strFileName { get; set; }
        public bool m_bHasKeyWord { get; set; }
        public string m_strLineMapping { get; set; }
        public FileExtension m_fileExtension { get; set; }

        public FileItem(string fileName, string strLineMapping = "", FileExtension fileExtension = default)
        {
            m_strFileName = fileName;
            m_strLineMapping = strLineMapping;
            m_fileExtension = fileExtension;
        }
    }
}
