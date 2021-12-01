using System.Collections.Generic;
using CommandLine;

namespace ModuleToDoc
{
    public sealed class CommandLineOptions
    {
        [Value(0, HelpText = "TripleCrown module folder", Required = true)]
        public string ModuleFolder { get; set; }
        
        [Option('o', "OutputFilename", HelpText = "Output filename")]
        public string OutputFile { get; set; }

        [Option('z', "ZonePivot", HelpText = "Zone Pivot version to use")]
        public string ZonePivot { get; set; }
    }
}