using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NDesk.Options;
using System.Text.RegularExpressions;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word       = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.IO;

namespace ConvertWithOffice
{
    class Program
    {
        static void Main(string[] args) {
            string out_format = "PDF";
            bool   optimize_for_screen = false;
            List<string> files = (new OptionSet{
                { "xps"   , v => { if (v != null) out_format = "XPS";         } },
                { "pdf"   , v => { if (v != null) out_format = "PDF";         } },
                { "screen", v => { if (v != null) optimize_for_screen = true; } }
            }).Parse(args);

            var rx_ppt = new Regex(@"\.(pptx|ppt|pptm|ppsx|pps|ppsm|potx|pot|potm|odp)$", RegexOptions.IgnoreCase);
            var rx_doc = new Regex(@"\.(docx|docm|dotx|dotm|doc|dot|htm|html|rtf|mht|mhtml|xml|odt)$", RegexOptions.IgnoreCase);

            PowerPoint.Application ppt = null;
            Word.Application       wrd = null;
            foreach (var i in files)
            {
                try
                {
                    var file = Path.GetFullPath(i);
                    if (rx_ppt.IsMatch(file))
                    {
                        if (ppt == null) ppt = new PowerPoint.Application();
                        PowerPoint.Presentation p = ppt.Presentations.Open(file, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                        string dst = getAvailableFileName(out_format == "PDF" ? getPDFName(file) : getXPSName(file));
                        p.SaveCopyAs(dst, out_format == "PDF" ? PowerPoint.PpSaveAsFileType.ppSaveAsPDF : PowerPoint.PpSaveAsFileType.ppSaveAsXPS, MsoTriState.msoTrue);
                        p.Close();
                    }
                    else if (rx_doc.IsMatch(file))
                    {
                        if (wrd == null) wrd = new Word.Application();
                        Word.Document d = wrd.Documents.Open(file, false, true);
                        string dst = getAvailableFileName(out_format == "PDF" ? getPDFName(file) : getXPSName(file));
                        d.ExportAsFixedFormat(dst, out_format == "PDF" ? Word.WdExportFormat.wdExportFormatPDF : Word.WdExportFormat.wdExportFormatXPS, false, optimize_for_screen ? Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen : Word.WdExportOptimizeFor.wdExportOptimizeForPrint);
                        d.Close(false);
                    }
                    else
                    {
                        Console.Error.WriteLine("Unknown file type: " + i);
                    }
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.Message);
                }
            }
        }

        private static string getPDFName (string file)
        {
            return Regex.Replace(file, @"\.[^.]+$", ".pdf");
        }

        private static string getXPSName(string file)
        {
            return Regex.Replace(file, @"\.[^.]+$", ".xps");
        }

        private static bool pathExists(string path)
        {
            return File.Exists(path) || Directory.Exists(path);
        }

        private static string getAvailableFileName(string file)
        {
            while (pathExists(file))
            {
                file = Regex.Replace(file, @"(?:\((\d+)\))?(\.[^.]+)$", m =>
                    m.Groups[1].Success
                        ? "(" + (int.Parse(m.Groups[1].Value) + 1) + ")" + m.Groups[2].Value
                        : " (1)" + m.Groups[2].Value
                );
            }
            return file;
        }
    }
}
