using System;
using System.IO;
using System.Threading.Tasks;
using CommandLine;
using DXPlus;

namespace ModuleToDoc
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            await new Program().Run(args);
        }

        private async Task Run(string[] args)
        {
            CommandLineOptions options = null;
            new Parser(cfg => { cfg.HelpWriter = Console.Error; })
                .ParseArguments<CommandLineOptions>(args)
                .WithParsed(clo => options = clo);
            if (options == null)
                return; // bad arguments or help.

            string outputFile = string.IsNullOrEmpty(options.OutputFile)
                ? Path.ChangeExtension(Path.GetFileName(options.ModuleFolder), ".docx")
                : options.OutputFile;

            if (File.Exists(outputFile))
            {
                await Console.Error.WriteLineAsync($"Output file {outputFile} already exists.");
                return;
            }
            
            var wordDocument = Document.Create(options.OutputFile);
            var processor = new ModuleProcessor(options.ModuleFolder);

            try
            {
                await processor.Process(wordDocument, options.ZonePivot);
                wordDocument.Save();
            }
            catch (Exception ex)
            {
                await Console.Error.WriteLineAsync(ex.Message);
                wordDocument.Close();
                File.Delete(outputFile);
            }
        }
    }
}