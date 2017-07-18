using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System.Drawing.Printing;
using System.Drawing;
using System.Windows.Forms;
using System.Printing;

namespace PhilipsLabelPrint
{
    class Program
    {
        static int rowOfMilliAmp = 2;
        static int columnOfMilliAmp = 13;

        static void Main(string[] args)
        {
            string myCurrent = getAmps();       //get current from philips csv file
            Console.WriteLine(myCurrent + "mA");     //output current to console

            //used to print current to the printer
            LocalPrintServer localPrintServer = new LocalPrintServer();
            PrintQueue defaultPrintQueue = LocalPrintServer.GetDefaultPrintQueue();

            PrintSystemJobInfo myPrintJob = defaultPrintQueue.AddJob();     //Call AddJob

            Stream myStream = myPrintJob.JobStream;     //Write a Byte buffer to the JobStream and close the stream
            Byte[] myByteBuffer = UnicodeEncoding.Unicode.GetBytes("The current of the driver is " + myCurrent + "mA");       //printed at the printer
            myStream.Write(myByteBuffer, 0, myByteBuffer.Length);
            myStream.Close();
        }

        //get Current from CSV file
        static string getAmps()
        {
            int count = 0;
            string amps = null;

            FileInfo newestFile = GetNewestFile(new DirectoryInfo(@"C:\Users\eabernathy\Desktop\SimpleSetFolder"));
            Console.WriteLine(newestFile);
            //gives the location for the csv file
            string filepath = @"C:\Users\eabernathy\Desktop\SimpleSetFolder\" + newestFile;

            using (var fs = File.OpenRead(filepath))
            using (StreamReader sr = new StreamReader(fs))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    var values = line.Split(',');

                    if (count == rowOfMilliAmp)
                    {
                        amps = values[columnOfMilliAmp];
                    }
                    count++;
                }
            }
            return amps;
        }
        public static FileInfo GetNewestFile(DirectoryInfo directory)
        {
            return directory.GetFiles()
                .Union(directory.GetDirectories().Select(d => GetNewestFile(d)))
                .OrderByDescending(f => (f == null ? DateTime.MinValue : f.LastWriteTime))
                .FirstOrDefault();
        }
    }
}