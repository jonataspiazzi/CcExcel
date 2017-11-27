using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace CcExcel.Test
{
    public static class DumpGeneratedExcelFiles
    {
        public static void Dump(Stream stream, [CallerMemberName] string callerName = null)
        {
            bool.TryParse(ConfigurationManager.AppSettings?.Get("DumpGeneratedExcelFiles"), out var isDumpActive);

            if (!isDumpActive) return;

            var date = new FileInfo(typeof(Excel).Assembly.Location).CreationTime;

            var dirName = Path.Combine(
                typeof(DumpGeneratedExcelFiles).Assembly.Location,
                $@"..\..\..\..\TestResults\ExcelFiles From Deploy {date:yyyy-MM-dd HH_mm_ss}");

            if (!Directory.Exists(dirName)) Directory.CreateDirectory(dirName);

            var className = new StackTrace().GetFrame(1).GetMethod().DeclaringType.Name;

            var fileName = Path.Combine(dirName, $"{className}_{callerName}.xlsx");

            using (var fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                stream.Position = 0;

                stream.CopyTo(fs);
            }
        }
    }
}
