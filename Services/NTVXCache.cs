using Microsoft.Office.Interop.Excel;
using NTVXApi.Models;
using System;
using System.Collections.Generic;
using System.Web;

namespace NTVXApi.Services
{
    public static class NTVXCache
    {
        public static void CacheData()
        {
            CacheClientVerionMapping();
        }

        private static void CacheClientVerionMapping()
        {
            using (ExcelWorkbook excel = new ExcelWorkbook())
            {
                string versionfile = HttpContext.Current.Server.MapPath("~/Excel/ClientVersions/Versions.xlsx");
                excel.Open(versionfile);

                Worksheet versionSheet = excel.OpenWorksheet("Version");
                int emptyCount = 0; // if there are two emptie A cells then exit
                int row = 1; // first row is heading
                
                while (true)
                {
                    row++;
                    ClientVersionModel model = new ClientVersionModel();
                    model.CurrentVersion = versionSheet.Cells[row, 2].Value;
                    if (string.IsNullOrEmpty(model.CurrentVersion))
                    {
                        emptyCount++;
                        if (emptyCount == 2)
                            break;
                        else
                            continue;
                    }

                    model.NewVersion = versionSheet.Cells[row, 3].Value;
                    model.VersionNum = Convert.ToInt32(versionSheet.Cells[row, 4].Value);
                    model.EndOfLife = Convert.ToBoolean(versionSheet.Cells[row, 5].Value);

                    if (!ClientVersion.VersionMapping.ContainsKey(model.CurrentVersion))
                    {
                        ClientVersion.VersionMapping.Add(model.CurrentVersion, new List<ClientVersionModel>());
                    }

                    ClientVersion.VersionMapping[model.CurrentVersion].Add(model);
                    emptyCount = 0;
                }
            }
        }
    }
}
