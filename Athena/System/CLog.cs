using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Athena
{
    class CLog
    {
        static public void WriteLog(long chatID, long senderKey, string name, string message, string command, string contents)
        {
            string chatName = "";

            if (chatID == -1001202203239)
                chatName = "클랜방";
            else if (chatID == -1001312491933)
                chatName = "운영진방";
            else if (chatID == -1001389956706)
                chatName = "사전안내방";
            else
                chatName = "Unknown";

           string year = DateTime.Now.Year.ToString("D4");
            string month = DateTime.Now.Month.ToString("D2");
            string day = DateTime.Now.Day.ToString("D2");
            string hour = DateTime.Now.Hour.ToString("D2");
            string min = DateTime.Now.Minute.ToString("D2");
            string second = DateTime.Now.Second.ToString("D2");

            string fileName = year + month + day + hour;
            string logTime = year + month + day + "-" + hour + min + second;

            string log = "[" + logTime + "-" + chatName + "-" + name + "(" + senderKey.ToString() + ")] " + "(Message : " + message + ") (Command : " + command + ") (Contents : " + contents + ")";
            File.AppendAllLines(@"Log/" + fileName + ".txt", new[] { log });
        }
    }
}
