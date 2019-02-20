using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace bmDataExtract
{
    public class BimetOneReader
    {
        public enum FileFormat
        {
            NoSession,
            NumeratedSessions,
            UnnumeratedSessions
        }
        public string FileName { get; private set; }
        public FileFormat Format { get; private set; }

        private string _fileContent;
        private Dictionary<int, decimal[]> _fileContentValues = new Dictionary<int, decimal[]>();


        public BimetOneReader(string fileName, FileFormat format)
        {
            FileName = fileName;
            Format = format;
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
                if (Format == FileFormat.NumeratedSessions)
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
                } else
                if (Format == FileFormat.UnnumeratedSessions)
                {
                    // Sessions
                    var sessions = _fileContent.Replace(".#IND", "").Replace(".#INF", "").Replace("\r\n", "\0").Split('\0');
                    int sessionId = 1;
                    foreach (var s in sessions)
                    {
                        if (!string.IsNullOrEmpty(s))
                        {
                            var values = s.Split(',');
                            _fileContentValues.Add(sessionId, values.Select(p => decimal.Parse(p, System.Globalization.NumberStyles.Float)).ToArray());
                            sessionId++;
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
                if (Format == FileFormat.NoSession) return new int[0];
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
