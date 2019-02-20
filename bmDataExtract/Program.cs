using bmDataExtract.Catalogs;
using CommandLine;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace bmDataExtract
{
    public class Options
    {
        [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
        public bool Verbose { get; set; }

        [Option('d', "directory", Required = true, HelpText = "Directory location")]
        public string Directory { get; set; }
        [Option('x', "excel", Required = true, HelpText = "Excel file")]
        public string Excel { get; set; }

        //
    }

    class Program
    {
        
        static void Main(string[] args)
        {
            Console.WriteLine("=====================================================");
            Console.WriteLine("==           Bimet Circus Extractor v1.1           ==");
            Console.WriteLine("=====================================================");
            Options options = new Options();
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       Console.WriteLine($"Current Arguments:");
                       Console.WriteLine($"Verbose: -v {o.Verbose}");
                       Console.WriteLine($"Directory: -d {o.Directory}");
                       Console.WriteLine($"Excel: -x {o.Excel}");
                       options = o;
                   });

            Console.WriteLine("=====================================================");
            Console.WriteLine("=========       MERIDIAN RELATIONSHIP      ==========");
            Console.WriteLine("=====================================================");
            MeridianHerarchy str = new MeridianHerarchy();
            for (int i= 1; i<=12; i++)
                Console.WriteLine($"Slave of {str.Meridians[i].Name} is " + str.SlaveOf(str.Meridians[i]).Name);
            Console.WriteLine("=====================================================");
            for (int i = 1; i <= 12; i++)
                Console.WriteLine($"Master of {str.Meridians[i].Name} is " + str.Meridians[i].Master.Name);
            Console.WriteLine("=====================================================");
            for (int i = 1; i <= 12; i++)
                Console.WriteLine($"Son of {str.Meridians[i].Name} is " + str.SonOf(str.Meridians[i]).Name);
            Console.WriteLine("=====================================================");
            for (int i = 1; i <= 12; i++)
                Console.WriteLine($"Mother of {str.Meridians[i].Name} is " + str.Meridians[i].Mother.Name);

            Console.WriteLine("=====================================================");
            Console.WriteLine("Press <ENTER> to Continue.");
            Console.ReadLine();
            
            ExtractBimet bm = new ExtractBimet(options.Directory, options.Excel);

            bm.Start();

            Console.WriteLine("Press <ENTER> to Exit.");
            Console.ReadLine();
        }
    }
}

