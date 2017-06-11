using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using HtmlAgilityPack;
using MoreLinq;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;

namespace SoundLibTool
{
	static class Downloader
	{
		public static void DownloadBrokenFiles()
		{
			var brokenFiles = Directory.EnumerateFiles(Paths.DownloadsPath, "*.part", SearchOption.TopDirectoryOnly).Select(x => Path.GetFileName(x.Substring(0, x.Length - 5)));
			var completedList = File.ReadAllLines(Paths.CompletedListPath).ToList();
			var downloadList = completedList.Where(x => brokenFiles.Contains(x.Split('|')[0])).ToList();
			completedList = completedList.Except(downloadList).ToList();
			//File.WriteAllLines(completedListFile, completedList);
			DownloadFiles(downloadList, completedList);
		}

		public static void DownloadFiles(string company)
		{
			var completedList = File.Exists(Paths.CompletedListPath) ? File.ReadAllLines(Paths.CompletedListPath).ToList() : new List<string>();
			var downloadList = File.ReadAllLines($@"{Paths.ListsPath}\{company}.txt").Except(completedList).ToList();
			DownloadFiles(downloadList, completedList);
		}

		static readonly Regex GoogleLinkIdRegEx = new Regex(@"https://drive.google.com/open\?id=(\w*)", RegexOptions.IgnoreCase);
		static readonly Regex GoogleLinkIdRegEx2 = new Regex(@"https://drive.google.com/(?:[\w.]*/)?file/d/(\w*)/view(?:\?usp=\w*)?", RegexOptions.IgnoreCase);

		static void DownloadFiles(IEnumerable<string> downloadList, ICollection<string> completedList)
		{
			FirefoxDriver driver = null;
			var initialPartCount = 0; //Directory.EnumerateFiles(Paths.DownloadsPath, "*.part").Count();
			foreach (var download in downloadList)
			{
				while (true)
				{
					var pendingDownloads = Directory.EnumerateFiles(Paths.DownloadsPath, "*.part").Count();
					if (pendingDownloads - initialPartCount < 5)
					{
						try
						{
							var destFile = Paths.DownloadsPath + @"\" + download.Split('|')[0];
							if (!File.Exists(destFile))
							{
								var destPartFile = destFile + ".part";
								if (File.Exists(destPartFile))
									File.Delete(destPartFile);
								if (driver == null)
								{
									driver = new FirefoxDriver(new FirefoxProfileManager().GetProfile("Test"));
									driver.Manage().Timeouts().PageLoad = new TimeSpan(0, 0, 0, 5);
									driver.Manage().Timeouts().ImplicitWait = new TimeSpan(0, 0, 0, 5);
									driver.Manage().Timeouts().AsynchronousJavaScript = new TimeSpan(0, 0, 0, 5);
								}
								var googleLink = download.Split('|')[1];
								var googleLinkMatch = GoogleLinkIdRegEx.Match(googleLink);
								if (googleLinkMatch.Groups.Count < 2)
									googleLinkMatch = GoogleLinkIdRegEx2.Match(googleLink);
								var googleLinkId = googleLinkMatch.Groups[1].Captures[0];
								var url = $@"https://drive.google.com/uc?id={googleLinkId}&export=download";
								driver.Navigate().GoToUrl(url);
								Thread.Sleep(1000);
								while (!File.Exists(destPartFile))
								{
									var elements = driver.FindElementsById("uc-download-link");
									if (elements.Any())
									{
										elements.First().Click();
										while (!File.Exists(destPartFile) || new FileInfo(destPartFile).Length == 0)
											Thread.Sleep(1000);
									}
									else Thread.Sleep(1000);
								}
							}
							if (!completedList.Contains(download))
							{
								completedList.Add(download);
								File.WriteAllLines(Paths.CompletedListPath, completedList);
							}
							break;
						}
						catch (NoSuchElementException)
						{
						}
						catch (WebDriverException e)
						{
							//driver?.Close();
							//driver = new FirefoxDriver(new FirefoxProfileManager().GetProfile("Test"));
						}
					}
					else Thread.Sleep(1000);
				}
			}
		}

		public static void GrabAllLinks()
		{
		    var profileManager = new FirefoxProfileManager();
            var profile = profileManager.GetProfile("Test");
            var firefoxService = FirefoxDriverService.CreateDefaultService();
		    var options = new FirefoxOptions() {Profile = profile};
		    var  driver = new FirefoxDriver(firefoxService, options, new TimeSpan(0, 0, 30));
            driver.Navigate().GoToUrl("http://audioclub.top/");
			var htmlDocument = new HtmlDocument();
			htmlDocument.LoadHtml(driver.PageSource);
			var elements = htmlDocument.DocumentNode.SelectNodesNoFail(".//li[contains(@class,'cat-item')]//a");
			var links = new HashSet<string>();
			foreach (var element in elements)
			{
				links.Add(Regex.Match(element.Attributes["href"].Value, @"http://audioclub.top/category/([^/]*)/").Groups[1].Captures[0].Value);
				Console.WriteLine("Found company {0}", element.InnerText);
			}
			links.Add("plugin");
			var contents = links.Aggregate("", (current, company) => current + GrabLinks(company, driver));
			File.WriteAllText($@"{Paths.ListsPath}\All.txt", contents);
		}

		public static string GrabLinks(string company, FirefoxDriver driver = null)
		{
			if (driver == null)
				driver = new FirefoxDriver(new FirefoxProfileManager().GetProfile("Test"));
		    var links = ExtractLinks(driver, $"http://hidelinks.in/{company}/");
            if (!links.Any())
				links = ExtractLinks(driver, $"http://hidelinks.in/{company}-gg/");
			if (!links.Any())
				Console.WriteLine($"No links were found for company {company}");
			var contents = links.Aggregate("", (current, link) => current + $"{link.InnerText}|{link.Attributes["href"].Value}\r\n");
		    if (contents == null) throw new ArgumentNullException(nameof(contents));
		    File.WriteAllText($@"{Paths.ListsPath}\{company}.txt", contents);
			return contents;
		}

		static List<HtmlNode> ExtractLinks(FirefoxDriver driver, string url )
	    {
	        driver.Navigate().GoToUrl(url);
	        var htmlDocument = new HtmlDocument();
	        htmlDocument.LoadHtml(driver.PageSource);
	        var links = htmlDocument.DocumentNode.SelectNodesNoFail(".//div[@class='entry-content']//a").Where(l => l.Attributes["href"].Value.Contains("drive.google.com")).ToList();
	        return links;
	    }

	    public static void RebuildCompletedList()
		{
			var files = Directory.EnumerateFiles(Paths.DownloadsPath, "*.rar", SearchOption.AllDirectories).Select(Path.GetFileName).ToHashSet(StringComparer.OrdinalIgnoreCase).ToList();
			files.Sort();
			var allKnownFilesContents = File.ReadAllLines(Paths.ListsPath + @"\_All.txt");
			var knownFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
			foreach (var line in allKnownFilesContents)
				knownFiles[line.Split('|')[0]] = line.Split('|')[1];
			var contents = "";
			foreach (var file in files)
			{
				string link;
				if (knownFiles.TryGetValue(file, out link))
					contents += $"{file}|{link}\r\n";
			}
			File.WriteAllText($@"{Paths.ListsPath}\_Completed.txt", contents);
		}
	}
}