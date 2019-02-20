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
        [Option('r', "relations", Required = false, HelpText = "Show Meridian Relationshipts")]
        public bool Relations { get; set; }

        //
    }

    class Program
    {
        static ILogger logger = new Logger();

        static void Main(string[] args)
        {
            logger.Log("================================================================", EventType.Info);
            logger.Log("==                Bimet Circus Extractor v1.1                 ==", EventType.Info);
            logger.Log("================================================================", EventType.Info);
            Options options = new Options();
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       logger.Log($"Current Arguments:");
                       logger.Log($"Verbose: -v {o.Verbose}");
                       logger.Log($"Directory: -d {o.Directory}");
                       logger.Log($"Excel: -x {o.Excel}");
                       logger.Log($"Relationships: -r");
                       options = o;
                   });

            if (options.Relations)
            {
                logger.Log("================================================================");
                logger.Log("=========           MERIDIAN RELATIONSHIP             ==========");
                logger.Log("================================================================");
                MeridianHerarchy str = new MeridianHerarchy();
                for (int i = 1; i <= 12; i++)
                    logger.Log($"Slave of {str.Meridians[i].Name} is " + str.SlaveOf(str.Meridians[i]).Name);
                logger.Log("================================================================");
                for (int i = 1; i <= 12; i++)
                    logger.Log($"Master of {str.Meridians[i].Name} is " + str.Meridians[i].Master.Name);
                logger.Log("================================================================");
                for (int i = 1; i <= 12; i++)
                    logger.Log($"Son of {str.Meridians[i].Name} is " + str.SonOf(str.Meridians[i]).Name);
                logger.Log("================================================================");
                for (int i = 1; i <= 12; i++)
                    logger.Log($"Mother of {str.Meridians[i].Name} is " + str.Meridians[i].Mother.Name);
                logger.Log("================================================================");
                logger.Log("================================================================");
                logger.Log("");
                logger.Log("Press <ENTER> to Continue.");
                Console.ReadLine();
            }

            ExtractBimet bm = new ExtractBimet(options.Directory, options.Excel, logger);

            bm.Start();

            logger.Log("Press <ENTER> to Exit.");
            Console.ReadLine();
        }
    }
}

