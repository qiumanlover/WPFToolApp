using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ChangeName
{
    class FileForRename
    {
        internal string OldFilePath { get; set; }
        internal string OldFileName { get; set; }
        public string NewFileName { get; set; }
        public string NewFilePath { get; set; }
        internal string FileDirPath { get; }
        internal string fileExt { get; }

        internal FileForRename(string filePath)
        {
            this.OldFilePath = filePath;
            this.OldFileName = Path.GetFileNameWithoutExtension(filePath);
            this.NewFileName = OldFileName;
            this.NewFilePath = OldFilePath;
            this.FileDirPath = Path.GetDirectoryName(filePath);
            this.fileExt = Path.GetExtension(filePath);
        }
        internal bool needRename()
        {
            return !OldFileName.Equals(NewFileName);
        }
        internal bool Rename()
        {
            File.Move(OldFilePath, this.NewFilePath);
            this.OldFileName = this.NewFileName;
            this.OldFilePath = this.NewFilePath;
            return true;
        }
        internal void ResetInfo()
        {
            this.NewFileName = this.OldFileName;
            this.NewFilePath = this.OldFilePath;
        }
        internal void ChangeName(string newFileName)
        {
            this.NewFileName = newFileName;
            this.NewFilePath = Path.Combine(this.FileDirPath, newFileName + this.fileExt);
        }
    }
}
