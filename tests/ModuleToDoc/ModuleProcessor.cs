using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DXPlus;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Renderer.Docx;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;
using MSLearnRepos;
using Constants = MSLearnRepos.Constants;

namespace ModuleToDoc
{
    public class ModuleProcessor
    {
        private readonly string moduleFolder;
        private readonly string indexFile;
        private TripleCrownModule moduleData;
        private IDocument document;
        private string zonePivot;

        public ModuleProcessor(string moduleFolder)
        {
            this.moduleFolder = moduleFolder;
            this.indexFile = Path.Combine(moduleFolder, Constants.IndexFile);
        }

        public async Task Process(IDocument wordDocument, string selectedZonePivot)
        {
            this.document = wordDocument ?? throw new ArgumentNullException(nameof(wordDocument));
            this.zonePivot = selectedZonePivot;
            
            if (!Directory.Exists(moduleFolder)
                || !File.Exists(indexFile))
            {
                throw new Exception($"No module found in folder {moduleFolder}");
            }

            var tcLoader = TripleCrownGitHubService.CreateLocal(moduleFolder);
            moduleData = TripleCrownModule.LoadFromFile(indexFile, string.Empty);
            await tcLoader.LoadUnitsAsync(moduleData);

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
            
            var docWriter = new DocxRenderer(wordDocument, moduleFolder);
            pipeline.Setup(docWriter);

            foreach (var unit in moduleData.Units.Skip(2).Take(1))
            {
                if (unit.Content != null)
                {
                    WriteUnitHeader(unit);
                    string markdownText = await tcLoader.ReadContentForUnitAsync(unit);
                    var markdownDocument = Markdown.Parse(markdownText, pipeline);
                    docWriter.Render(markdownDocument);
                }
            }
        }

        private void WriteUnitHeader(TripleCrownUnit unit)
        {
            document.AddPageBreak();
            document.AddParagraph(unit.Title).Style(HeadingType.Heading1);
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