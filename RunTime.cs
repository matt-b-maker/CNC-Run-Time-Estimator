using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFSWTry
{
    class RunTime
    { 
        public RunTime (string fileName, int seconds)
        {
            this.FileName = fileName;
            this.Seconds = seconds;
        }

        public string FileName { get; set; }

        public int Seconds { get; set; }

        public string IndividualTimeOutput(int seconds)
        {
            TimeSpan t = TimeSpan.FromSeconds(seconds);
            string output = t.ToString(@"hh\:mm\:ss");
            return output;
        }

        public string ShorterFileName(string fileName)
        {
            string[] realName = fileName.Split('\\');
            return realName[realName.Length - 1];
        }
    }
}
