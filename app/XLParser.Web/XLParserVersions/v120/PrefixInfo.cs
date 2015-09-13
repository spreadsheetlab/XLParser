using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLParser.Web.XLParserVersions.v120
{
    public class PrefixInfo
    {
        
        public string FilePath { get; }
        public bool HasFilePath => FilePath != null;

        private readonly int? fileNumber;
        public int FileNumber => fileNumber.Value;
        public bool HasFileNumber => fileNumber.HasValue;

        public string FileName { get; }
        public bool HasFileName => FileName != null;

        public bool HasFile => HasFileName || HasFileNumber;

        public string Sheet { get; }
        public bool HasSheet => Sheet != null;

        public string MultipleSheets { get; }
        public bool HasMultipleSheets => MultipleSheets != null;

        public bool IsQuoted { get; }

        public PrefixInfo(string sheet = null, int? fileNumber = null, string fileName = null, string filePath = null, string multipleSheets = null, bool isQuoted = false)
        {
            Sheet = sheet;
            this.fileNumber = fileNumber;
            FileName = fileName;
            FilePath = filePath;
            MultipleSheets = multipleSheets;
            IsQuoted = isQuoted;
        }

        public override string ToString()
        {
            string res = "";
            if (IsQuoted) res += "'";
            if (HasFilePath) res += FilePath;
            if (HasFileNumber) res += $"[{FileNumber}]";
            if (HasFileName) res += $"[{FileName}]";
            if (HasSheet) res += Sheet;
            if (HasMultipleSheets) res += MultipleSheets;
            if (IsQuoted) res += "'";
            res += "!";
            return res;
        }

    }
}
