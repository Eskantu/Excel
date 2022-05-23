using ExcelFile;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelFile excel = new ExcelFile();
            //string[] headers = new string[] { "Nombre", "Edad", "Direccion" };
            //string[] data = new string[] { "Mario,23,Conocido", "Mario,23,Conocido", "Mario,23,Conocido", "Mario,23,Conocido", "Mario,23,Conocido", };
            //excel.WriteIntoFile(headers, data, @"C:\Users\maescalante\Desktop\Archivo.xlsx", "Hoja 1");

            excel.CreateFile(@"C:\Users\maescalante\Desktop\hojas.csv", @"C:\Users\maescalante\Desktop\Headers.csv", @"C:\Users\maescalante\Desktop\CXCCReporteEstatusConvenioRD.csv");
        }
    }
}
