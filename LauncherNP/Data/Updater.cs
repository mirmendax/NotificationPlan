using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using LauncherNP.Models;
using Newtonsoft.Json;
using Settings = LauncherNP.Models.Settings;

namespace LauncherNP.Data
{
    public class Updater
    {
        /// <summary>
        /// Проверка наличия обновлений
        /// </summary>
        /// <returns>bool, UpdateFileInfo. Если true то UpdateFileInfo заполнен из файла обновления</returns>
        public static (bool, UpdateFileInfo) IsEnableUpdate()
        {
            SetLauncherContext sContext = new SetLauncherContext();
            var result = false;
            UpdateFileInfo ufInfo = new UpdateFileInfo();
            if (Directory.Exists(sContext.Settings.PathToUpdate))
            {
                var filePath = sContext.Settings.PathToUpdate + sContext.Settings.FileInfo;
                if (File.Exists(filePath))
                {
                    
                    var json = File.ReadAllText(filePath);
                    ufInfo = JsonConvert.DeserializeObject<UpdateFileInfo>(json);
                    if (ufInfo.Build > sContext.Settings.Build)
                    {
                        result = true;
                    }
                }
            }
            return (result, ufInfo);
        }

        public static void Update(string path)
        {
            var sContext = new SetLauncherContext();
            (bool, UpdateFileInfo) isEnableUpdate = IsEnableUpdate();
            if (isEnableUpdate.Item1)
            {
                var updateFileInfo = isEnableUpdate.Item2;
                try
                {
                    var updateDirectory = path + "tempupdate";
                    File.Copy(sContext.Settings.PathToUpdate + updateFileInfo.UpdateFileName, path + "temp.zip", true);
                    
                    Directory.CreateDirectory(updateDirectory);
                    ZipFile.ExtractToDirectory(path + "temp.zip", updateDirectory);
                    
                    ProcessStartInfo prcUpdate = new ProcessStartInfo(updateDirectory + "\\path.exe");
                    prcUpdate.WorkingDirectory = Path.GetDirectoryName(prcUpdate.FileName) ?? throw new Exception();
                    var prc = new Process();
                    prc.StartInfo = prcUpdate;
                    prc.Start();
                    prc.WaitForExit();
                    
                    Directory.Delete(updateDirectory,true);
                    File.Delete(path + "temp.zip");
                    sContext.Settings.Build = updateFileInfo.Build;
                    sContext.SaveSettings();
                }
                catch (Exception )
                {
                    
                }
                
            }
            
        }
    }
}