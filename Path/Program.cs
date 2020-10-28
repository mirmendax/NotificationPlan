using System;
using System.IO;
using System.Net.Mime;
using System.Threading;

namespace Path
{
    internal class Program
    {

        public static void Main(string[] args)
        {
            
            Console.WriteLine("Начало установки.");
            var path = Directory.GetCurrentDirectory();
            Console.WriteLine(path);
            var dirApp = System.IO.Path.GetDirectoryName(path);
            var files = Directory.GetFiles(path);
            
            Console.WriteLine("Копирование файлов");
            foreach (var file in files)
            {
                if (file.Contains("Path.exe")) continue;
                if (System.IO.Path.GetFileName(file) == "") continue;
                Console.WriteLine($"Копирование файла: {file}");

                File.Copy(file, dirApp + "\\" + System.IO.Path.GetFileName(file), true);
                Thread.Sleep(400);
            }

            Console.WriteLine("Установка закончена.");
            Thread.Sleep(1500);
        }
    }
}