using System;
using static FileKeywordSearcher.Form1;

namespace FileKeywordSearcher
{
    internal class FileItem
    {
        public string m_strFileName { get; set; }
        public bool m_bHasKeyWord { get; set; }
        public string m_strLineMapping { get; set; }

        public FileItem(string fileName, bool hasKeyword = false, string strLineMapping = "")
        {
            m_strFileName = fileName;
            m_bHasKeyWord = hasKeyword;
            m_strLineMapping = strLineMapping;
        }
    }
}
