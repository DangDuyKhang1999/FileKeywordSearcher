using System;

namespace FileKeywordSearcher
{
    internal class FileItem
    {
        public string m_strFileName { get; set; }
        public bool m_bHasKeyWord { get; set; }

        public FileItem(string fileName, bool hasKeyword)
        {
            m_strFileName = fileName;
            m_bHasKeyWord = hasKeyword;
        }
    }
}
