using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace bmDataExtract
{
    public class BimetOneReader
    {
        public string FileName { get; private set; }
        public bool haveSessions { get; private set; }

        private string _fileContent;
        private Dictionary<int, decimal[]> _fileContentValues = new Dictionary<int, decimal[]>();


        public BimetOneReader(string fileName, bool havesessions)
        {
            FileName = fileName;
            haveSessions = havesessions;
            ReadFile();
        }

        public void Refresh()
        {
            ReadFile();
        }

        private void ReadFile()
        {

            _fileContent = File.ReadAllText(FileName);
            try
            {
                if (haveSessions)
                {
                    // Sessions
                    var sessions = _fileContent.Replace(".#IND", "").Replace(".#INF", "").Replace("\r\n", "\0").Split('\0');
                    foreach (var s in sessions)
                    {
                        if (!string.IsNullOrEmpty(s))
                        {
                            var values = s.Split(',');
                            if (!_fileContentValues.ContainsKey(int.Parse(values[0])))
                            _fileContentValues.Add(int.Parse(values[0]), values.Skip(1).Select(p => decimal.Parse(p, System.Globalization.NumberStyles.Float)).ToArray());
                        }
                    }
                }
                else
                {
                    string[] values = _fileContent.Split(',');
                    
                        _fileContentValues.Add(0, values.Select(p => decimal.Parse(p, System.Globalization.NumberStyles.Float)).ToArray());
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            if (_fileContentValues.Keys.Count == 0) throw new ApplicationException("File incomplete");
        }

        public int[] Sessions
        {
            get
            {
                if (_fileContent == null) ReadFile();
                if (!haveSessions) return new int[0];
                return _fileContentValues.Keys.ToArray();
            }
        }

        public decimal this[int i]
        {
            get
            {
                if (_fileContent == null) ReadFile();
                return _fileContentValues[0][i];
            }
            set { _fileContentValues[0][i] = value; }
        }

        public decimal this[int session, int i]
        {
            get {
                if (_fileContent == null) ReadFile();
                return _fileContentValues[session][i];
            }
            set { _fileContentValues[session][i] = value; }
        }
    }
}
