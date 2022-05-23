using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace ExcelFile
{
    public class ExcelFile
    {
        /// <summary>
        /// Crea un archivo de excel con el <paramref name="path"/> dado
        /// </summary>
        /// <param name="headers">Encabezados de la tabla</param>
        /// <param name="data">Datos de tabla separados por comas</param>
        /// <param name="path">Ruta donde se creara el archivo</param>
        /// <returns>ruta donde se encuentra el archivo</returns>
        public string CreateFile(string[] headers, string[] data, string path = "")
        {
            string fileTemp = Path.Combine(Directory.GetCurrentDirectory(), "temp.xlsx");
            string pathFileFinal = string.IsNullOrEmpty(path) ? Path.Combine(Directory.GetCurrentDirectory(), "Archivo.xlsx") : path;
            string sheetName = "Hoja 1";
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileTemp, SpreadsheetDocumentType.Workbook))
            {
                CreateSheet(sheetName, package);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                byte[] byteArray = File.ReadAllBytes(fileTemp);
                ms.Write(byteArray, 0, byteArray.Length);
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(ms, true))
                {

                    UpdateExcelFile(document, headers, data, sheetName);
                }
                File.WriteAllBytes(pathFileFinal, ms.ToArray());
            }
            File.Delete(fileTemp);
            return pathFileFinal;


        }

        /// <summary>
        /// Escribe en un archivo ya creado la informacion dada
        /// </summary>
        /// <param name="headers">Headers de la tabla a escribir</param>
        /// <param name="data">Datos a escribir separados por comas</param>
        /// <param name="path">Ruta de archivo donde se escribiran los datos</param>
        /// <param name="sheetName">Nombre de la hoja donde donde se escribiran los datos</param>
        /// <returns>la ruta del archivo actualizado</returns>
        public string WriteIntoFile(string[] headers, string[] data, string path, string sheetName)
        {
            FileInfo fileInfo = new FileInfo(path);
            string fileTemp = Path.Combine(Directory.GetCurrentDirectory(), $"{fileInfo.Name}");
            File.Copy(path, fileTemp);

            using (MemoryStream ms = new MemoryStream())
            {
                byte[] byteArray = File.ReadAllBytes(fileTemp);
                ms.Write(byteArray, 0, byteArray.Length);
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(ms, true))
                {

                    UpdateExcelFile(document, headers, data, sheetName);
                }
                File.WriteAllBytes(path, ms.ToArray());
            }
            File.Delete(fileTemp);
            return path;


        }


        /// <summary>
        /// Crea un archivo de excel en el <paramref name="filePath"/> con hojas y datos dados por archivos .CSV
        /// </summary>
        /// <param name="sheetpath">Path de archvo .CSV con el nombre de las hojas a crear</param>
        /// <param name="headerPath">Path de archvo .CSV con los headers de la tabla a crear</param>
        /// <param name="dataPath">Path de archvo .CSV con los datos a insertar</param>
        /// <param name="filePath">Path donde se creara el archivo de excel</param>
        /// <returns>path del archvio craeado</returns>
        public string CreateFile(string sheetpath, string headerPath, string dataPath, string filePath = "")
        {
            //Creacion de archivo y hojas
            string fileTemp = Path.Combine(Directory.GetCurrentDirectory(), "temp.xlsx");
            string pathFileFinal = string.IsNullOrEmpty(filePath) ? Path.Combine(Directory.GetCurrentDirectory(), "Archivo.xlsx") : filePath;
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileTemp, SpreadsheetDocumentType.Workbook))
            {
                var sheets = ReadSheets(sheetpath).ToArray();
                    CreateSheet(sheets, package);
            }
            string[] headers = ReadHeaders(headerPath).ToArray();
            string[] data = ReadData(dataPath).ToArray();
            string archivo = "";
            foreach (var item in ReadSheets(sheetpath))
            {
                archivo = WriteIntoFile(headers, data, filePath, item);
            }
            File.Copy(archivo, pathFileFinal);
            File.Delete(archivo);
            return pathFileFinal;
        }

        private static void CreateSheet(string sheetName, SpreadsheetDocument package)
        {
            UInt32 sheetId = 1;
            WorkbookPart workbookPart = package.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            package.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());


            var sheetpart = package.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            sheetpart.Worksheet = new Worksheet(sheetData);
            Sheets sheets = package.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationShipId = package.WorkbookPart.GetIdOfPart(sheetpart);
            int totalSheets = sheets.Elements<Sheet>().Count();
            if (totalSheets > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }
            Sheet sheet = new Sheet() { Id = relationShipId, SheetId = sheetId, Name = $"{sheetName}" };
            sheets.Append(sheet);
        }

        private static void CreateSheet(string[] sheetsArray, SpreadsheetDocument package)
        {
            UInt32 sheetId = 1;
            WorkbookPart workbookPart = package.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            package.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            var sheetpart = package.WorkbookPart.AddNewPart<WorksheetPart>();

            foreach (var item in sheetsArray)
            {
                var sheetData = new SheetData();
                sheetpart.Worksheet = new Worksheet(sheetData);
                Sheets sheets = package.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationShipId = package.WorkbookPart.GetIdOfPart(sheetpart);
                int totalSheets = sheets.Elements<Sheet>().Count();
                if (totalSheets > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }
                Sheet sheet = new Sheet() { Id = relationShipId, SheetId = sheetId, Name = $"{item}" };
                sheets.Append(sheet);
            }

        }

        private void UpdateExcelFile(SpreadsheetDocument document, string[] headers, string[] data, string sheetName)
        {
            WorksheetPart worksheetPart = GetWorkSheetpartByname(document, sheetName);

            if (worksheetPart != null)
            {
                // Create new Worksheet
                Worksheet worksheet = new Worksheet();
                worksheetPart.Worksheet = worksheet;

                // Create new SheetData
                SheetData sheetData = new SheetData();

                Row tRowHeader = new Row();
                foreach (var header in headers)
                {
                    tRowHeader.Append(CreateCell(header));
                }
                sheetData.Append(tRowHeader);

                foreach (var item in data)
                {
                    Row tRow = new Row();
                    foreach (var cellData in item.Split(','))
                    {
                        tRow.Append(CreateCell(cellData));
                    }
                    sheetData.Append(tRow);
                }
                worksheet.Append(sheetData);

                worksheetPart.Worksheet.Save();
            }
            document.WorkbookPart.Workbook.Save();
        }

        private WorksheetPart GetWorkSheetpartByname(SpreadsheetDocument document, string sheetName)
        {
            var hojas = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
                       Elements<Sheet>();
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
                       Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                return null;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        private Cell CreateCell(string text) => new Cell
        {
            DataType = ResolveCellDataTypeOnValue(text),
            CellValue = new CellValue(text)
        };

        private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text) =>
            int.TryParse(text, out _) || double.TryParse(text, out _) ? CellValues.Number : CellValues.String;



        #region MetodosParaLecturaDeArchivos

        private static IEnumerable<string> ReadSheets(string sheetsPath)
        {
            using (StreamReader reader = new StreamReader(sheetsPath))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine().Split(',');
                    foreach (var item in line)
                    {
                        yield return item;
                    }
                }
            }
        }
        private static IEnumerable<string> ReadHeaders(string headersPath)
        {
            using (StreamReader reader = new StreamReader(headersPath))
            {
                while (!reader.EndOfStream)
                {
                    yield return reader.ReadLine();
                }
            }
        }

        private static IEnumerable<string> ReadData(string dataPath, string filtro = "")
        {
            using (StreamReader reader = new StreamReader(dataPath))
            {
                while (!reader.EndOfStream)
                {
                    yield return reader.ReadLine();
                }
            }
        }

        #endregion
    }
}
