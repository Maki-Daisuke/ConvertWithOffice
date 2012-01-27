using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.IO;

namespace ConvertWithPowerPoint
{
    class Program
    {
        static void Main(string[] args) {
            foreach (var i in args) {
                Console.Out.WriteLine(i);
            }
            try {
                PowerPoint.Application ppt = new PowerPoint.Application();
                Console.Out.WriteLine("1");
                PowerPoint.Presentation p = ppt.Presentations.Open(Path.GetFullPath(args[0]), MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                Console.Out.WriteLine("2");
                p.SaveCopyAs(Path.GetFullPath(args[1]), PowerPoint.PpSaveAsFileType.ppSaveAsPDF, MsoTriState.msoTrue);
                Console.Out.WriteLine("3");
                p.Close();
                Console.Out.WriteLine("4");
                ppt.Quit();
            } catch (COMException e) {
                Console.Out.WriteLine(e.Message);
            }
        }
    }
}
