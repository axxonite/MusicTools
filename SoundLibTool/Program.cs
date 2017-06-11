using System.Linq;
using MoreLinq;

namespace SoundLibTool
{
	static class Program
	{
		static void Main(string[] args)
		{
			Enumerable.Range(0, args.Length).ForEach(i => args[i] = args[i].TrimEnd('\\'));
			switch (args[0])
			{
				case "-x":
				{
					Installer.InstallLibs();
					break;
				}
				case "-dl":
				{
					Downloader.GrabLinks(args[1]);
					break;
				}
				case "-dal":    
				{
					Downloader.GrabAllLinks();
					break;
				}
				case "-df":
				{
					Downloader.DownloadFiles(args[1]);
					break;
				}
				case "-db":
				{
					Downloader.DownloadBrokenFiles();
					break;
				}
				case "-rdl":
				{
					Downloader.RebuildCompletedList();
					break;
				}
				case "-fl":
				{
					Installer.FixLibraryFolders();
					break;
				}
				case "-cl":
				{
					Installer.CategorizeLibraryFolders();
					break;
				}
				case "-pi":
				{
					Installer.ProcessIsos();
					break;
				}
			}
		}
	}
}