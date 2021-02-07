using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace PresentationToJpgs2
{
    class Program
    {
        static void Main(string[] args)
        {
            string datatime = Convert.ToString(DateTime.Now);
            datatime = datatime.Replace(" ", "");
            datatime = datatime.Replace(".", "");
            datatime = datatime.Replace(":", "");
            datatime = Convert.ToString(Convert.ToInt64(datatime));

            string[] presentations = new DirectoryInfo(Directory.GetCurrentDirectory()).GetFiles("*.pptx").Select(f => f.FullName).ToArray();
        
            for (int i = 0; i < presentations.Length; i++)
            {
                PowerPoint.Application PPApp = new PowerPoint.Application();

                File.Move($"{presentations[i]}", $"{datatime}.pptx");

                PowerPoint.Presentation Pres = PPApp.Presentations.Open(Directory.GetCurrentDirectory() + $"\\{datatime}.pptx", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

                Pres.SaveCopyAs((Directory.GetCurrentDirectory() + $"\\{Pres.Name}"), PowerPoint.PpSaveAsFileType.ppSaveAsJPG);

                Pres.Application.Quit();

                datatime = Convert.ToString(Convert.ToInt64(datatime) + 1);
            }
        }
    }
}
