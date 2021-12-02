using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MSLearnRepos;

namespace ModuleToDoc
{
    public sealed class LearnUtilities
    {
        private Action<string> _logger;
        private string _accessToken;

        public async Task<(TripleCrownModule module, string markdownFile)> DownloadModuleAsync(
            ITripleCrownGitHubService tcService, string accessToken,
            string learnFolder, string outputFolder, Action<string> logger = null)
        {
            if (string.IsNullOrEmpty(outputFolder))
                throw new ArgumentException($"'{nameof(outputFolder)}' cannot be null or empty.", nameof(outputFolder));

            _accessToken = string.IsNullOrEmpty(accessToken)
                ? GithubHelper.ReadDefaultSecurityToken()
                : accessToken;

            _logger = logger ?? Console.WriteLine;

            var module = await tcService.GetModuleAsync(learnFolder);
            if (module == null)
                throw new ArgumentException($"Failed to parse Learn module from {learnFolder}", nameof(learnFolder));

            await tcService.LoadUnitsAsync(module);

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            _logger?.Invoke($"Copying \"{module.Title}\" to {outputFolder}");

            var markdownFile = Path.Combine(outputFolder, Path.ChangeExtension(Path.GetFileNameWithoutExtension(learnFolder),".md"));

            await using var tempFile = new StreamWriter(markdownFile);

            foreach (var unit in module.Units)
            {
                await tempFile.WriteLineAsync($"# {unit.Title}");
                var text = await tcService.ReadContentForUnitAsync(unit);
                if (text != null)
                {
                    text = await DownloadAllImagesForUnit(text, tcService, learnFolder, outputFolder);
                    await tempFile.WriteLineAsync(text);
                    await tempFile.WriteLineAsync();
                }

                if (unit.Quiz != null)
                {
                    if (text == null)
                        await tempFile.WriteLineAsync($"## {unit.Quiz.Title}\r\n");

                    foreach (var question in unit.Quiz.Questions)
                    {
                        await tempFile.WriteLineAsync($"### {question.Content}");
                        foreach (var choice in question.Choices)
                        {
                            await tempFile.WriteAsync(choice.IsCorrect ? "- [X] " : "- [ ]");
                            await tempFile.WriteLineAsync(choice.Content);
                            await tempFile.WriteLineAsync();
                        }
                    }

                    await tempFile.WriteLineAsync();
                }
            }

            return (module, markdownFile);
        }

        private async Task<string> DownloadAllImagesForUnit(string markdownText, ITripleCrownGitHubService gitHub, string moduleFolder, string tempFolder)
        {
            var images = Regex.Matches(markdownText, @"!\[.*\]\((.*?)\)")
                .Union(Regex.Matches(markdownText, @"<img.+src=(?:\""|\')(.+?)(?:\""|\')(?:.+?)\>"))
                .Union(Regex.Matches(markdownText, @":::image.+source=(?:\""|\')(.+?)(?:\""|\')(?:.+?):::"))
                .ToList();

            foreach (Match match in images)
            {
                string imagePath = match.Groups[1].Value;
                if (string.IsNullOrEmpty(imagePath)) continue;

                string newPath = await DownloadImageAsync(imagePath, gitHub, moduleFolder, tempFolder);
                if (newPath != null)
                {
                    markdownText = markdownText.Replace(imagePath, newPath);
                }
            }

            return markdownText;
        }

        private async Task<string> DownloadImageAsync(string imagePath, ITripleCrownGitHubService gitHub, string moduleFolder, string tempFolder)
        {
            // Ignore urls.
            if (imagePath.StartsWith("http"))
                return null;

            // Remove any relative path info
            if (imagePath.StartsWith(@"../") || imagePath.StartsWith(@"..\"))
                imagePath = imagePath[3..];

            string remotePath = moduleFolder + "/" + imagePath;
            string localPath = Path.Combine(tempFolder, imagePath);

            // Already downloaded?
            if (File.Exists(localPath))
                return imagePath;

            string localFolder = Path.GetDirectoryName(localPath);
            if (!string.IsNullOrEmpty(localFolder))
            {
                if (!Directory.Exists(localFolder))
                    Directory.CreateDirectory(localFolder);
            }

            try
            {
                var (binary, _) = await gitHub.ReadFileForPathAsync(remotePath);
                if (binary != null)
                {
                    await File.WriteAllBytesAsync(localPath, binary);
                }
                else
                {
                    throw new Exception($"{remotePath} did not return an image as expected.");
                }
            }
            catch (Octokit.ForbiddenException)
            {
                // Image > 1Mb in size, switch to the Git Data API and download based on the sha.
                var remote = (IRemoteTripleCrownGitHubService)gitHub;
                await GitHelper.GetAndWriteBlobAsync(Constants.Organization,
                    gitHub.Repository, remotePath, localPath, _accessToken, remote.Branch);
            }

            return imagePath;
        }

        public static async Task<(string repo, string branch, string folder)> RetrieveLearnLocationFromUrlAsync(string moduleUrl)
        {
            using var client = new HttpClient();
            string html = await client.GetStringAsync(moduleUrl);

            string pageKind = Regex.Match(html, @"<meta name=""page_kind"" content=""(.*?)""\s/>").Groups[1].Value;
            if (pageKind != "module")
                throw new ArgumentException("URL does not identify a Learn module - use the module landing page URL", nameof(moduleUrl));

            string lastCommit = Regex.Match(html, @"<meta name=""original_content_git_url"" content=""(.*?)""\s/>").Groups[1].Value;
            var uri = new Uri(lastCommit);
            if (uri.Host.ToLower() != "github.com")
                throw new ArgumentException("Identified module not hosted on GitHub", nameof(moduleUrl));

            var path = uri.LocalPath.ToLower().Split('/').Where(s => !string.IsNullOrWhiteSpace(s)).ToList();

            if (path[0] != "microsoftdocs")
                throw new ArgumentException("Identified module not in MicrosoftDocs organization", nameof(moduleUrl));

            string repo = path[1];
            if (!repo.StartsWith("learn-"))
                throw new ArgumentException("Identified module not in recognized MS Learn GitHub repo", nameof(moduleUrl));

            if (path.Last() == "index.yml")
                path.RemoveAt(path.Count - 1);

            string branch = path[3];
            string folder = string.Join('/', path.Skip(4));

            return (repo, branch, folder);
        }
    }
}
