using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Net.WebRequestMethods;
using File = System.IO.File;

namespace Davinci_Delivery_Exporter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string InstallLocation = "";
            string DeliveryPresets = "";

            bool GetShortcutTarget(string shortcut)
            {
                WshShell shell = new WshShell();
                IWshShortcut link = (IWshShortcut)shell.CreateShortcut(shortcut);
                if (link.TargetPath.ToLower().Contains("resolve.exe")) if (File.Exists(Path.GetDirectoryName(link.TargetPath) + "\\DataBase\\Resolve Projects\\Settings\\DeliverPresetList.xml"))
                    {
                        InstallLocation = Path.GetDirectoryName(link.TargetPath);
                        DeliveryPresets = Path.GetDirectoryName(link.TargetPath) + "\\DataBase\\Resolve Projects\\Settings\\DeliverPresetList.xml";
                    }
                return true;
            }

            foreach (string file in Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop))) if (file.EndsWith(".lnk")) GetShortcutTarget(file);
            if (string.IsNullOrWhiteSpace(InstallLocation))
            {
                foreach (string file in Directory.GetFiles("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs")) if (file.EndsWith(".lnk")) GetShortcutTarget(file);
            }
            if (string.IsNullOrWhiteSpace(InstallLocation))
            {
                foreach (string dirs in Directory.GetDirectories("C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs"))
                {
                    foreach (string file in Directory.GetFiles(dirs)) if (file.EndsWith(".lnk")) GetShortcutTarget(file);
                    foreach (string dirssecond in Directory.GetDirectories(dirs))
                    {
                        foreach (string file in Directory.GetFiles(dirssecond)) if (file.EndsWith(".lnk")) GetShortcutTarget(file);
                    }
                }
            }
            if (string.IsNullOrWhiteSpace(InstallLocation))
            {
                foreach (string file in Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs")) if (file.EndsWith(".lnk")) GetShortcutTarget(file);
            }
            if (string.IsNullOrWhiteSpace(InstallLocation))
            {
                foreach (string dirs in Directory.GetDirectories(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs"))
                {
                    foreach (string file in Directory.GetFiles(dirs)) if (file.EndsWith(".lnk")) GetShortcutTarget(file);
                    foreach (string dirssecond in Directory.GetDirectories(dirs))
                    {
                        foreach (string file in Directory.GetFiles(dirssecond)) if (file.EndsWith(".lnk")) GetShortcutTarget(file);
                    }
                }
            }

            bool ListPresets()
            {
                StreamReader reader = new StreamReader(DeliveryPresets);
                string PresetsFileData = reader.ReadToEnd();
                reader.Close();
                reader.Dispose();

                string listData = "";
                int spaceCounter;

                int CurrentLevel = 0;

                foreach (string line in PresetsFileData.Replace("\r\n", "\n").Split('\n'))
                {
                    if (line.Contains("<DbKey>")) CurrentLevel++;
                    if (CurrentLevel > 0)
                    {
                        listData = listData + line + "\n";
                    }
                    if (line.Contains("</DbKey>")) CurrentLevel--;
                }

                spaceCounter = listData.Split('<').First().Length;
                foreach (string line in listData.Split('\n'))
                {
                    if (line.Split('<').First().Length == spaceCounter) Console.WriteLine(line.Split('>')[1].Split('<')[0]);
                }

                return true;
            }

            bool SendImportData(string InputFile)
            {
                bool FileExists = true;
                if (InputFile.Contains(":")) if (!File.Exists(InputFile)) FileExists = false;

                if (!FileExists)
                {
                    Console.WriteLine("error: input file does not exist.\n");
                }
                else
                {
                    string oldData = "";
                    string outputData = "";
                    bool endPresetList = false;

                    StreamReader reader = new StreamReader(DeliveryPresets);
                    oldData = reader.ReadToEnd();
                    reader.Close();
                    reader.Dispose();

                    foreach (string line in oldData.Replace("\r\n", "\n").Split('\n'))
                    {
                        if (line.Contains("</PresetList>"))
                        {
                            if (endPresetList == false)
                            {
                                StreamReader Inputreader = new StreamReader(InputFile);
                                outputData = outputData + Inputreader.ReadToEnd() + "\r\n";
                                Inputreader.Close();
                                Inputreader.Dispose();
                                endPresetList = true;
                            }
                        }

                        outputData = outputData + line + "\r\n";
                    }

                    StreamWriter writer = new StreamWriter(DeliveryPresets);
                    writer.Write(outputData.Substring(0, outputData.Length - 2));
                    writer.Close();
                    writer.Dispose();

                    Console.WriteLine("done!\n");
                }
                return true;
            }

            bool GetExportData(string PresetName, string OutputFile)
            {
                bool DirectoryExists = true;
                if (OutputFile.Contains(":")) if (!Directory.Exists(Path.GetDirectoryName(OutputFile))) DirectoryExists = false;

                if (DirectoryExists)
                {
                    StreamReader reader = new StreamReader(DeliveryPresets);
                    string PresetsFileData = reader.ReadToEnd();
                    string ExportData = "";

                    bool copyData = false;

                    foreach (string line in PresetsFileData.Replace("\r\n", "\n").Split('\n'))
                    {
                        if (copyData == true && !line.Contains("</PresetList>"))
                        {
                            if (line.Contains("<SyRecordInfo DbId=")) ExportData = ExportData + line.Split('"').First() + "\"123456789\">" + "\r\n";
                            else if (line.Contains("<FieldsBlob>")) ExportData = ExportData + line.Split('>').First() + ">0000000000</FieldsBlob>" + "\r\n";
                            else ExportData = ExportData + line + "\r\n";
                        }
                        if (line.Contains("<PresetList>")) copyData = true;
                        if (line.Contains("</PresetList>")) copyData = false;
                    }

                    bool FoundPreset = true;

                    if (!string.IsNullOrWhiteSpace(PresetName))
                    {
                        FoundPreset = false;
                        PresetsFileData = ExportData;
                        ExportData = "";

                        PresetsFileData = PresetsFileData.Replace("\r\n", "\n");

                        int CurrentLevel = 0;
                        foreach (string line in PresetsFileData.Split('\n'))
                        {
                            if (line.Contains("<Element>")) CurrentLevel++;
                            if (line.Contains("</Element>")) CurrentLevel--;
                            if (CurrentLevel > 0)
                            {
                                if (line.Contains("<SyRecordInfo DbId=")) ExportData = ExportData + line.Split('"').First() + "\"123456789\">" + "\r\n";
                                else if (line.Contains("<FieldsBlob>")) ExportData = ExportData + line.Split('>').First() + ">0000000000</FieldsBlob>" + "\r\n";
                                else ExportData = ExportData + line + "\r\n";
                            }
                            else
                            {
                                ExportData = ExportData + line + "\r\n";
                                if (ExportData.Contains(PresetName))
                                {
                                    FoundPreset = true;
                                    break;
                                }
                                else ExportData = "";
                            }
                        }
                    }

                    reader.Close();
                    reader.Dispose();

                    if (FoundPreset == true)
                    {
                        if (!string.IsNullOrWhiteSpace(ExportData))
                        {
                            StreamWriter writer = new StreamWriter(OutputFile);
                            writer.Write(ExportData.Substring(0, ExportData.Length - 2));
                            writer.Close();
                            writer.Dispose();
                        }
                        Console.WriteLine("done!\n");
                    }
                    else Console.WriteLine("error: failed to find preset. Please check spelling as this is case sensitive\n");
                    return true;
                }
                else
                {
                    Console.WriteLine("error: output path does not exist.\n");
                    return false;
                }
            }


            string WelcomeMessage = "" +
                        "----------------------------------------------------------------------\n" +
                        "       Davinci Delivery Exporter - Created by Ricky Divjakovski\n" +
                        "----------------------------------------------------------------------\n";

            string ShowHelp = "\n" +
                        "Usage: \n" +
                        "-export [PresetName] [OutputFile]    | Exports single delivery preset\n" +
                        "-import [InputFile]                  | Imports all delivery presets\n" +
                        "                                     |\n" +
                        "-list                                | Lists all available presets\n" +
                        "                                     |\n" +
                        "-help                                | Displays this help message\n" +
                        "-version                             | Displays version info";

            if (string.IsNullOrWhiteSpace(InstallLocation))
            {
                Console.WriteLine("Error: failed to find davinci install path.\n\nExiting..");
            }
            else
            {
                Console.Write(WelcomeMessage);
                if (args.Length > 1)
                {
                    if (args[0] == "-export")
                    {
                        if (args.Length == 2)
                        {
                            Console.Write("\nExporting delivery presets.. ");
                            GetExportData("", args[1]);
                        }
                        else if (args.Length == 3)
                        {
                            Console.Write("\nExporting delivery preset " + args[1] + ".. ");
                            GetExportData(args[1], args[2]);
                        }
                    }
                    if (args[0] == "-import")
                    {
                        if (args.Length == 2)
                        {
                            Console.Write("\nImporting delivery presets.. ");
                            SendImportData(args[1]);
                        }
                    }
                }
                else if (args.Length == 1)
                {
                    if (args[0] == "-list")
                    {
                        Console.WriteLine("\nPresets available for export-");
                        ListPresets();
                    }
                    if (args[0] == "-help")
                    {
                        Console.WriteLine(ShowHelp);
                    }
                    if (args[0] == "-version")
                    {
                        var ver = Assembly.GetExecutingAssembly().GetName().Version;
                        string version = string.Format("{0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision, Assembly.GetEntryAssembly().GetName().Name);
                        Console.WriteLine("\nversion - " + version);
                    }
                }
                else Console.WriteLine(ShowHelp);
            }
        }
    }
}
