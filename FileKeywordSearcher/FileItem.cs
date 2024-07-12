using System;
using static FileKeywordSearcher.Form1;

namespace FileKeywordSearcher
{
    internal class FileItem
    {
        public string m_strFileName { get; set; }
        public bool m_bHasMultiKeyWord { get; set; }
        public string m_strLineMapping { get; set; }
        public eFileExtension m_fileExtension { get; set; }

        public FileItem(string fileName, string strLineMapping = "", eFileExtension fileExtension = default, bool bHasMultiKeyWord = false)
        {
            m_strFileName = fileName;
            m_strLineMapping = strLineMapping;
            m_fileExtension = fileExtension;
            m_bHasMultiKeyWord = bHasMultiKeyWord;
        }
    }
}
