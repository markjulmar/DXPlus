using CommandLine;

namespace ModuleToDoc
{
    public sealed class CommandLineOptions
    {
        [Option('i', "input", Required = true)]
        public string InputFolder { get; set; }

        [Option('o', "OutputFilename", HelpText = "Output filename")]
        public string OutputFile { get; set; }

        [Option('r', "Repo")]
        public string GitHubRepo { get; set; }

        [Option('b', "Branch")]
        public string GitHubBranch { get; set; }

        [Option('t', "Token")]
        public string AccessToken { get; set; }

        [Option('z', "ZonePivot", HelpText = "Zone Pivot version to use")]
        public string ZonePivot { get; set; }

        [Option('d', "Debug", HelpText = "Display Markdown graph (DEBUG)")]
        public bool Debug { get; set; }
    }
}