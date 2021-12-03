using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DXPlus;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Extensions.Tables;
using Markdig.Renderer.Docx;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;
using MSLearnRepos;
using Octokit;

namespace ModuleToDoc
{
    public class ModuleProcessor
    {
        private TripleCrownModule moduleData;
        private IDocument document;
        private string zonePivot;
        private string markdownFile;
        private readonly ITripleCrownGitHubService tcService;
        private readonly string accessToken;
        private readonly string moduleFolder;

        public static async Task<ModuleProcessor> CreateFromUrl(string url, string accessToken = null)
        {
            if (string.IsNullOrEmpty(url))
                throw new ArgumentNullException(nameof(url));

            var (repo, branch, folder) = await LearnUtilities.RetrieveLearnLocationFromUrlAsync(url);
            return await CreateFromRepo(repo, branch, folder, accessToken);
        }

        public static Task<ModuleProcessor> CreateFromRepo(string repo, string branch, string folder, string accessToken = null)
        {
            if (string.IsNullOrEmpty(repo))
                throw new ArgumentException($"'{nameof(repo)}' cannot be null or empty.", nameof(repo));
            if (string.IsNullOrEmpty(branch))
                throw new ArgumentException($"'{nameof(branch)}' cannot be null or empty.", nameof(branch));
            if (string.IsNullOrEmpty(folder))
                throw new ArgumentException($"'{nameof(folder)}' cannot be null or empty.", nameof(folder));

            accessToken = string.IsNullOrEmpty(accessToken)
                ? GithubHelper.ReadDefaultSecurityToken()
                : accessToken;

            var tcService = TripleCrownGitHubService.CreateFromToken(repo, branch, accessToken);
            return Task.FromResult(new ModuleProcessor(tcService, accessToken, folder));
        }

        public static Task<ModuleProcessor> CreateFromLocalFolder(string learnFolder)
        {
            if (string.IsNullOrWhiteSpace(learnFolder))
                throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

            if (!Directory.Exists(learnFolder))
                throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

            var tcService = TripleCrownGitHubService.CreateLocal(learnFolder);
            return Task.FromResult(new ModuleProcessor(tcService, null, learnFolder));
        }

        private ModuleProcessor(ITripleCrownGitHubService tcService, string accessToken, string moduleFolder)
        {
            this.tcService = tcService ?? throw new ArgumentNullException(nameof(tcService));
            this.moduleFolder = moduleFolder ?? throw new ArgumentNullException(nameof(moduleFolder));
            this.accessToken = accessToken;
        }

        public async Task Process(IDocument wordDocument, string selectedZonePivot)
        {
            this.document = wordDocument ?? throw new ArgumentNullException(nameof(wordDocument));
            this.zonePivot = selectedZonePivot;

            string outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), Path.GetFileNameWithoutExtension(Path.GetTempFileName()));
            (moduleData, markdownFile) = await new LearnUtilities().DownloadModuleAsync(tcService, accessToken, moduleFolder, outputFolder);

            try
            {
                AddMetadata();
                WriteTitle();

                var context = new MarkdownContext();
                var pipelineBuilder = new MarkdownPipelineBuilder();
                var pipeline = pipelineBuilder
                    .UseAdvancedExtensions()
                    .UseEmphasisExtras(EmphasisExtraOptions.Strikethrough)
                    .UseIncludeFile(context)
                    .UseQuoteSectionNote(context)
                    .UseRow(context)
                    .UseTripleColon(context)
                    .Build();

                var docWriter = new DocxObjectRenderer(wordDocument, moduleFolder);

                string markdownText = File.ReadAllText(markdownFile);
                MarkdownDocument markdownDocument = Markdown.Parse(markdownText, pipeline);

                //MarkdigDebug.Dump(markdownDocument);

                docWriter.Render(markdownDocument);
            }
            finally
            {
                if (outputFolder != null)
                {
                    Directory.Delete(outputFolder, true);
                }
            }
        }

        private void AddMetadata()
        {
            document.SetPropertyValue(DocumentPropertyName.Title, moduleData.Title);
            document.SetPropertyValue(DocumentPropertyName.Subject, moduleData.Summary);
            document.SetPropertyValue(DocumentPropertyName.CreatedDate, moduleData.LastUpdated.ToString("yyyy-MM-ddT00:00:00Z"));
            document.SetPropertyValue(DocumentPropertyName.Creator, moduleData.Metadata.MsAuthor);
        }

        private void WriteTitle()
        {
            document.AddParagraph(moduleData.Title)
                .Style(HeadingType.Title);
            document.AddParagraph($"Last modified on {moduleData.LastUpdated.ToShortDateString()}")
                .Style(HeadingType.Subtitle);
        }
    }
}