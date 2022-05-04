using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReporter
{
    public class WorkFileInfo
    {
        public Guid Id
        {
            get
            {
                using (MD5 md5 = MD5.Create())
                {
                    byte[] hash = md5.ComputeHash(Encoding.Default.GetBytes(this.FilePath.ToLower()));
                    Guid result = new Guid(hash);

                    return result;
                }
            }
        }

        public string FilePath
        {
            get; set;
        }

        public string SheetName
        {
            get; set;
        }

        public string[] SheetNames
        {
            get; set;
        }

        public int HeaderRowIndex
        {
            get; set;
        } = -1;

        public HeaderField[] HeaderLabels
        {
            get; set;
        } = null;

        public string GetFileName()
        {
            return Path.GetFileName(this.FilePath);
        }
    }
}