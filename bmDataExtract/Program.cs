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

            
            OrgansHerarchy str = new OrgansHerarchy();
             
            ExtractBimet bm = new ExtractBimet(options.Directory, options.Excel);

            bm.Start();

            Console.WriteLine("Press <enter> to finish.");
            Console.ReadLine();
        }
    }
}

