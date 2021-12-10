using System;
using System.IO;
using System.Threading.Tasks;
using CommandLine;
using DXPlus;

namespace ModuleToDoc
{
    public static class Program
    {
        public static async Task Main(string[] args)
        {
            CommandLineOptions options = null;
            new Parser(cfg => { cfg.HelpWriter = Console.Error; })
                .ParseArguments<CommandLineOptions>(args)
                .WithParsed(clo => options = clo);
            if (options == null)
                return; // bad arguments or help.

            if (string.IsNullOrEmpty(options.OutputFile))
                options.OutputFile = Path.GetFileName(options.InputFolder);
            options.OutputFile = Path.ChangeExtension(options.OutputFile, ".docx");

            if (string.IsNullOrEmpty(options.InputFolder))
            {
                Console.WriteLine("Missing input folder or URL.");
                return;
            }

            if (string.IsNullOrEmpty(options.OutputFile))
            {
                Console.WriteLine("Missing output file.");
                return;
            }

            if (File.Exists(options.OutputFile))
                File.Delete(options.OutputFile);

            ModuleProcessor processor = null;

            if (options.InputFolder.StartsWith("http", StringComparison.InvariantCultureIgnoreCase))
            {
                processor = await ModuleProcessor.CreateFromUrl(options.InputFolder, options.AccessToken);
            }
            else if (!string.IsNullOrEmpty(options.GitHubRepo))
            {
                if (string.IsNullOrEmpty(options.GitHubBranch))
                    options.GitHubBranch = "live";
                processor = await ModuleProcessor.CreateFromRepo(options.GitHubRepo, options.GitHubBranch, options.InputFolder, options.AccessToken);
            }
            else if (Directory.Exists(options.InputFolder))
            {
                processor = await ModuleProcessor.CreateFromLocalFolder(options.InputFolder);
            }

            if (processor == null)
            {
                await Console.Error.WriteLineAsync("Please supply a Url, local folder, or GitHub details to a Learn module.");
                return;
            }

            using var wordDocument = Document.Create(options.OutputFile);
            await processor.Process(wordDocument);
            wordDocument.Save();

            Console.WriteLine("Done.");
        }
    }
}