using System;
using System.IO;
using System.IO.Compression;

namespace Ashok.UiPath.OpenOfficeInterop
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            string originalPath = @"C:\Users\akalbhor\Documents\Open Office Interops\Sample.ott";
            string tempPath = @"C:\Users\akalbhor\Documents\Open Office Interops\Sample.zip";
            string extractPath = @"C:\Users\akalbhor\Documents\Open Office Interops\Extracted";

            if (File.Exists(@"C:\Users\akalbhor\Documents\Open Office Interops\Sample.ott"))
            {
                File.Move(originalPath.Trim(), tempPath.Trim());
                ZipFile.ExtractToDirectory(tempPath, extractPath);
            }
            int a = 10;

        }
    }
}
