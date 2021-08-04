using Microsoft.Office.Interop.Excel;
using NTVXApi.Models;
using System;
using System.Collections.Generic;
using System.Web;

namespace NTVXApi.Services
{
    public static class WebTemplateCache
    {
        public static void CacheData()
        {
            CacheWebTemplates();
        }

        public static List<WebTemplate> Templates { get; } = new List<WebTemplate>();

        private static void CacheWebTemplates()
        {
            using (ExcelWorkbook excel = new ExcelWorkbook())
            {
                string versionfile = HttpContext.Current.Server.MapPath("~/Excel/WebTemplates/WebTemplates.xlsx");
                excel.Open(versionfile);

                Worksheet templateSheet = excel.OpenWorksheet("Templates");
                int emptyCount = 0; // if there are two empty A cells then exit
                int row = 1; // first row is heading

                string templateId;
                WebTemplate model;
                while (true)
                {
                    row++;

                    //reading web templates
                    templateId = Convert.ToString(templateSheet.Cells[row, 1].Value);

                    if (string.IsNullOrEmpty(templateId))
                    {
                        emptyCount++;
                        if (emptyCount == 2)
                            break;
                        else
                            continue;
                    }

                    model = Templates.Find(t => t.TemplateId == Convert.ToInt32(templateId));
                    if(model == null)
                    {
                        model = new WebTemplate();
                        model.TemplateId = Convert.ToInt32(templateId);
                        model.TemplateFile = templateSheet.Cells[row, 2].Value;
                        model.ServerTemplate = templateSheet.Cells[row, 3].Value;
                        model.DeleteFile = templateSheet.Cells[row, 9].Value;
                        model.Actions = new List<WebAction>();

                        Templates.Add(model);
                    }

                    //actions
                    model.Actions.Add(new WebAction()
                    {
                        Button = templateSheet.Cells[row, 4].Value,
                        UnlockRanges = templateSheet.Cells[row, 5].Value,
                        CopySheetsAndRanges = templateSheet.Cells[row, 6].Value,
                        Macro = templateSheet.Cells[row, 7].Value,
                        PasteSheetsAndRanges = templateSheet.Cells[row, 8].Value,
                    });

                    emptyCount = 0;
                }
            }
        }
    }
}