using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelExtractor.XML;

namespace ExcelExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("WELCOME TO EXCEL EXTRACTOR!");
            Console.WriteLine("---------------------------");

            if (File.Exists("example.xml"))
                try
                {
                    new Serializing("example.xml").Do();
                }
                catch (Exception E)
                {
                    Console.Error.WriteLine(E.ToString());
                }
            else foreach (var arg in args)
            {
                try
                {
                    new Serializing(arg).Do();
                }
                catch (Exception E)
                {
                    Console.Error.WriteLine(E.ToString());
                }
                
            }

            Console.WriteLine("----------------------------------");
            Console.WriteLine("EXCEL EXTRACTOR IS GOING TO SLEEP!");

        }
    }
}
