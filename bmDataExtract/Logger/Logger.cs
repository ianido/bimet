using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace bmDataExtract
{
    public class Logger : ILogger
    {
        public void Log(string message, EventType type = EventType.None, bool finishLine = true, bool includeDate = false)
        {
            Console.ResetColor();
            switch (type)
            {
                case EventType.Error: { Console.ForegroundColor = ConsoleColor.DarkRed; } break;
                case EventType.Warning: { Console.ForegroundColor = ConsoleColor.DarkYellow; } break;
                case EventType.Info: { Console.ForegroundColor = ConsoleColor.Cyan; } break;
                case EventType.Success: { Console.ForegroundColor = ConsoleColor.Green; } break;
            }
            if (string.IsNullOrEmpty(message) || message == @"\*" || message == @"*/") { Console.WriteLine(""); return; };
            if (finishLine)
                Console.WriteLine((includeDate ? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " : "") + message);
            else
                Console.Write((includeDate ? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " : "") + message);
            Console.ResetColor();
        }
    }
}
