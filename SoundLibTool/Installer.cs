using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using MoreLinq;
using Directory = Pri.LongPath.Directory;
using DirectoryInfo = Pri.LongPath.DirectoryInfo;
using File = Pri.LongPath.File;
using FileInfo = Pri.LongPath.FileInfo;
using Path = Pri.LongPath.Path;

namespace SoundLibTool
{
	static class Installer
	{
		public static void InstallLibs()
		{
			// Unmount();
			var files = Directory.EnumerateFiles(Paths.ToProcessPath, "*").ToArray();
			var archives = GetArchiveFiles(files);
			foreach (var archive in archives)
			{
				CleanupTempFolders();
				try
				{
					Console.Write("Processing {0}\r\n", archive);
					ProcessArchive(Paths.ToProcessPath, Paths.ToProcessPath, archive, files, Paths.OutputPath);
					var associatedArchives = GetAssociatedArchives(archive, files);
					foreach (var file in associatedArchives)
						Directory.Move(file, file.Replace(Paths.ToProcessPath, Paths.ProcessedPath));
					Console.WriteLine("Successfully processed {0}", archive);
				}
				catch (Exception e)
				{
					foreach (var key in e.Data.Keys)
						Console.WriteLine(e.Data[key] + ": " + key);
					Console.Write(e.Message);
					Console.WriteLine("Errors while processing {0}", archive);
					foreach (var file in GetAssociatedArchives(archive, files))
						Directory.Move(file, Paths.ToProcessPath + @"\_Broken\" + Path.GetFileName(file));
				}
				finally
				{
					Console.WriteLine();
					CleanupTempFolders();
				}
			}
		}

		static void CleanupTempFolders()
		{
			if (Directory.Exists(Paths.TransferPath))
			{
				MarkFolderWritable(Paths.TransferPath);
				Directory.Delete(Paths.TransferPath, true);
			}
			if (Directory.Exists(Paths.RamDriveTransferPath))
			{
				MarkFolderWritable(Paths.RamDriveTransferPath);
				Directory.Delete(Paths.RamDriveTransferPath, true);
			}
		}

		static void CopyFiles(string source, string dest, bool move, bool includeArchives = true)
		{
			if ( move )
				MarkFolderWritable(source);
			MarkFolderWritable(dest);
			Console.WriteLine("{0} {1} => {2}", move ? "Moving" : "Copying", source, dest);
			source = source.TrimEnd('\\');
			dest = dest.TrimEnd('\\');
			if ( !Directory.Exists(dest))
				Directory.CreateDirectory(dest);
			Directory.EnumerateDirectories(source, "*", SearchOption.AllDirectories).Select(d => d.Replace(source, dest)).ForEach(path => Directory.CreateDirectory(path));
			foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories).Where(f => Path.GetExtension(f) != ".nfo" && !Regex.IsMatch(Path.GetFileName(f), "All.Collection.Upload|WareZ-Audio", RegexOptions.IgnoreCase)).ToArray())
			{
				if (Path.GetExtension(file) == ".sfv")
					continue;
				//if (!includeArchives && Regex.IsMatch(Path.GetExtension(file), @"\.(rar|r\d+|zip|iso)"))
				if (!includeArchives && Path.GetDirectoryName(file) == source && Regex.IsMatch(Path.GetExtension(file), @"\.(rar|r\d+|zip)"))
						continue;
				var newFile = file.Replace(source, dest);
				if (move)
				{
					if (File.Exists(newFile))
						File.Delete(newFile);
					File.Move(file, newFile);
				}
				else File.Copy(file, newFile, true);
			}
		}

		static string EscapeSpecialChars(string contents)
		{
			var specialChars = new[] { "\\", ".", "$", "^", "{", "[", "(", "|", ")", "*", "+", "?" };
			var result = contents;
			specialChars.ForEach(c => result = result.Replace(c, @"\" + c));
			return result;
		}

		static IEnumerable<string> GetArchiveFiles(string[] files)
		{
			var results = new List<string>();
			foreach (var file in files)
			{
				if (/*Path.GetExtension(file) == ".iso" || */ Path.GetExtension(file) == ".zip") // Any iso or zip file.
					results.Add(file);
                //if (Path.GetExtension(file) == ".nrg")
                //    results.Add(file);
                if (Path.GetExtension(file) == ".r00" && files.All(f => String.Compare(f, file.Substring(0, file.Length - 4) + ".rar", StringComparison.OrdinalIgnoreCase) != 0)) // Any .r00 file that doesn't have a corresponding .rar file.
					results.Add(file);
				if (Path.GetExtension(file) == ".rar" && (!Regex.IsMatch(file, @"(?i)part[0-9]+\.rar$") || Regex.IsMatch(file, @"(?i)part0*1\.rar$"))) // Any rar file doesn't end with partxxx.rar, unless its the first part.
					results.Add(file);
			}
			return results.ToArray();
		}

		static IEnumerable<string> GetAssociatedArchives(string file, IEnumerable<string> files)
		{
			string pattern;
			if (file.EndsWith(".part1.rar"))
				pattern = EscapeSpecialChars(file.Substring(0, file.Length - 10)) + @"\.part\d+.rar";
			else if (file.EndsWith(".part01.rar"))
				pattern = EscapeSpecialChars(file.Substring(0, file.Length - 11)) + @"\.part\d+.rar";
			else if (Path.GetExtension(file) == ".r00")
				pattern = EscapeSpecialChars(file.Substring(0, file.Length - 4)) + @"\.r\d+";
			else return new[] { file };
			return files.Where(f => Regex.IsMatch(f, pattern, RegexOptions.IgnoreCase));
		}

		static string GetTempPath(long size)
		{
			return Paths.TransferPath;
			//var ramDriveSpaceAvaiable = new DriveInfo(Path.GetPathRoot(Paths.RamDriveTransferPath)).AvailableFreeSpace - 1000000000;
			//return size < ramDriveSpaceAvaiable ? Paths.RamDriveTransferPath : Paths.TransferPath;
		}

		static void MarkFolderWritable(string folder)
		{
			// ReSharper disable once ObjectCreationAsStatement
			if (!Directory.Exists(folder))
				return;
			Directory.GetDirectories(folder, "*", SearchOption.AllDirectories).ForEach(f => new DirectoryInfo(f) { Attributes = FileAttributes.Normal });
			Directory.GetFiles(folder, "*", SearchOption.AllDirectories).ForEach(f => new FileInfo(f) { Attributes = FileAttributes.Normal });
		}

		static void ProcessArchive(string rootPath, string folder, string file, IEnumerable<string> files, string destFolder)
		{
			if (Regex.IsMatch(file, @"(?i)\.(rar|r00|zip)$"))
			{
				var zipSize = GetAssociatedArchives(file, files).Sum(archiveFile => new FileInfo(archiveFile).Length);
				var tempPath = GetTempPath(zipSize * 2);
				var zipContentsPath = tempPath + folder.Substring(rootPath.Length) + Path.GetFileName(file);
				Console.WriteLine("Extracting {0} => {1}", Path.GetFileName(file), zipContentsPath);
				if (StartProcess(@"C:\Program Files\WinRAR\WinRar.exe", $"x -o+ -ibck -inul -y \"{file}\" \"{zipContentsPath}\"\\") < 2 && Directory.Exists(zipContentsPath))
				{
					MarkFolderWritable(zipContentsPath);
					ProcessFolder(tempPath, zipContentsPath, destFolder, true, false);
				}
				else
				{
					Console.WriteLine("Error while extracting {0}", file);
					throw new UnzipError();
				}
				Directory.Delete(zipContentsPath, true);
			}
			/*
			else if (Path.GetExtension(file) == ".iso")
			{
				Console.WriteLine("Mounting {0} => G:", file);
				StartProcess(@"C:\Program Files\PowerISO\piso.exe", string.Format("mount \"{0}\" G:", file));
				var isoSize = new FileInfo(file).Length;
				var tempPath = GetTempPath(isoSize);
				var isoContentsPath = tempPath + folder.Substring(rootPath.Length);
				CopyFiles("G:", isoContentsPath, false);
				MarkFolderWritable(isoContentsPath);
				ProcessFolder(tempPath, isoContentsPath, destFolder, true, true, string.Format(@"\{0}", Path.GetFileName(file)));
				Unmount();
				Directory.Delete(isoContentsPath, true);
			}
			* */
		}

		static void ProcessFolder(string rootPath, string baseFolder, string destFolder, bool canMove, bool preservePaths, string extraPath = "")
		{
			var folder = baseFolder;
			var files = Directory.EnumerateFiles(folder, "*").Where(f => Path.GetExtension(f) != ".txt" && Path.GetExtension(f) != ".nfo" && Path.GetExtension(f) != ".sfv").ToArray();
			while (!preservePaths)	// Not a true while, preserve paths does not change.
			{
				var subFolders = Directory.EnumerateDirectories(folder).ToArray();
				if (files.Length == 0 && subFolders.Length == 1)
					folder = subFolders[0];
				else break;
				files = Directory.EnumerateFiles(folder, "*").Where(f => Path.GetExtension(f) != ".txt" && Path.GetExtension(f) != ".nfo" && Path.GetExtension(f) != ".sfv").ToArray();
			}

			//files = Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories).ToArray();
			var archiveFiles = GetArchiveFiles(files);
			archiveFiles.ForEach(archive => ProcessArchive(rootPath, Path.GetDirectoryName(archive), archive, files, destFolder));

			var names = folder.Substring(rootPath.Length).Split(Path.DirectorySeparatorChar);
			var name = names.FirstOrDefault(n => Regex.IsMatch(n, "KONTAKT", RegexOptions.IgnoreCase)) ?? names[0];
			//if ((folder + "\\").StartsWith(string.Format("{0}\\{1}\\", baseFolder, name)))
			//	name = "";
			//name = Regex.Replace(name, @"\.(rar|r00|zip|iso)$", "", RegexOptions.IgnoreCase);
			name = Regex.Replace(name, @"\.(rar|r00|zip)$", "", RegexOptions.IgnoreCase);
			name = Regex.Replace(name, @"\.part1", "", RegexOptions.IgnoreCase);
			name = Regex.Replace(name, @"\.part01", "", RegexOptions.IgnoreCase);
			name = Regex.Replace(name, "-", " ");
			name = Regex.Replace(name, @"([a-z ])\.+([\w ])", "$1 $2", RegexOptions.IgnoreCase);	// Remove dots but leave those that are between digits.
			name = Regex.Replace(name, @"([\w ])\.+([a-z ])", "$1 $2", RegexOptions.IgnoreCase);
			name = Regex.Replace(name, @"\b(HYBRID|DYNAMiCS|PROPER|VSTi|RTAS|CHAOS|AMPLiFY|AU|MATRiX|DVDR|WAV|AiR|ArCADE|VR|CDDA|PAD|MiDi|CoBaLT|DiSCOVER)\b", "");
			name = Regex.Replace(name, @"\b(WareZ Audio info|Kontakt|Audiostrike|SYNTHiC4TE|AUDIOXiMiK|MAGNETRiXX|TZ7iSO|KLI|DVDriSO|DVD9|KRock|ACiD|REX|RMX|SynthX|AiFF|Apple Loops|AiRISO|MULTiFORMAT|AudioP2P|GHOSTiSO|REX2|DXi|HYBRiD|AKAI|ALFiSO)\b", "", RegexOptions.IgnoreCase);
			//name = Regex.Replace(name, "(HYBRID|DYNAMiCS|PROPER|VSTi|RTAS|CHAOS|AMPLiFY| AU |MATRiX)", "");
			//name = Regex.Replace(name, "(WareZ Audio info|Kontakt|Audiostrike|SYNTHiC4TE|AUDIOXiMiK||MAGNETRiXX)", "", RegexOptions.IgnoreCase);
			name = Regex.Replace(name, @"  +", " ");
			name = name.Trim(' ', '-');
			var destPath = $"{destFolder}\\{name}{extraPath}";
			CopyFiles(folder, destPath, canMove, false);
		}

		static int StartProcess(string exe, string arguments)
		{
			var process = new Process { StartInfo = new ProcessStartInfo { CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardInput = true, UseShellExecute = false, FileName = exe, Arguments = arguments } };
			//process.OutputDataReceived += ((sender, e) => Console.Write(e.Data + "\n"));
			process.Start();
			process.BeginOutputReadLine();
			process.WaitForExit();
			return process.ExitCode;
		}

		//static void Unmount()
		//{
		//	Console.WriteLine("Unmounting G:");
		//	StartProcess(@"C:\Program Files\PowerISO\piso.exe", "unmount G:");
		//}

		public static void CategorizeLibraryFolders()
		{
			var folders = Directory.EnumerateDirectories(Paths.OutputPath, "*", SearchOption.TopDirectoryOnly);
			foreach (var folder in folders)
			{
				if (Path.GetFileName(folder) == "_Kontakt" || Path.GetFileName(folder) == "_Loops")
					continue;
				var files = Directory.EnumerateFiles(folder, "*.nki", SearchOption.AllDirectories).ToArray();
				if (files.Length > 0)
					Directory.Move(folder, Paths.OutputPath + @"\_Kontakt\" + Path.GetFileName(folder));
				else
				{
					files = Directory.EnumerateFiles(folder, "*.iso", SearchOption.AllDirectories).ToArray();
					files = files.Concat(Directory.EnumerateFiles(folder, "*.bin", SearchOption.AllDirectories)).ToArray();
					if (files.Length > 0)
						Directory.Move(folder, Paths.OutputPath + @"\_Iso\" + Path.GetFileName(folder));
					else Directory.Move(folder, Paths.OutputPath + @"\_Loops\" + Path.GetFileName(folder));
				}
			}
		}

		public static void ProcessIsos()
		{
			var files = Directory.EnumerateFiles(Paths.OutputPath, "*.iso", SearchOption.AllDirectories).Concat(Directory.EnumerateFiles(Paths.OutputPath, "*.bin", SearchOption.AllDirectories));
			//var files = Directory.EnumerateFiles(Paths.OutputPath, "*.bin", SearchOption.AllDirectories);
			foreach (var file in files)
			{
				Console.WriteLine("Mounting {0} => D:", file);
				if (StartProcess(@"C:\Program Files\PowerISO\piso.exe", $"mount \"{file}\" D:") >= 0)
				{
					var destPath = Path.GetDirectoryName(file) + @"\" + Path.GetFileNameWithoutExtension(file) + @"\";
					CopyFiles("D:", destPath, false);					
				}
			}
		}


		public static void FixLibraryFolders()
		{
			var folders = Directory.EnumerateDirectories(Paths.OutputPath, "*", SearchOption.TopDirectoryOnly);
			foreach (var folder in folders)
			{
				var name = Path.GetFileName(folder);
				var oldName = name;
				name = Regex.Replace(name, @"\b(HYBRID|DYNAMiCS|PROPER|VSTi|RTAS|CHAOS|AMPLiFY|AU|MATRiX|DVDR|WAV|AiR|ArCADE|VR|CDDA|PAD|MiDi|CoBaLT|DiSCOVER)\b", "");
				name = Regex.Replace(name, @"\b(WareZ Audio info|Kontakt|Audiostrike|SYNTHiC4TE|AUDIOXiMiK|MAGNETRiXX|TZ7iSO|KLI|DVDriSO|DVD9|KRock|ACiD|REX|RMX|SynthX|AiFF|Apple Loops|AiRISO|MULTiFORMAT|AudioP2P|GHOSTiSO|REX2|DXi|HYBRiD|AKAI|ALFiSO)\b", "", RegexOptions.IgnoreCase);
				name = Regex.Replace(name, @"  +", " ");
				if (name != oldName && !Directory.Exists(Path.GetDirectoryName(folder) + @"\" + name))
					File.Move(folder, Path.GetDirectoryName(folder) + @"\" + name);
			}
		}

		class UnzipError : Exception
		{
		};
	}
}
