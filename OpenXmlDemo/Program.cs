using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;

namespace OpenXmlDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var stream = new MemoryStream())
            {
                var document = SpreadsheetDocument.Create(stream: stream,
                                                          type: SpreadsheetDocumentType.Workbook,
                                                          autoSave: true);

                var workbookPart = document.AddWorkbookPart();
                var workbook = workbookPart.Workbook = new Workbook();
                var sheetsElement = new Sheets();
                workbook.AppendChild(sheetsElement);

                InsertSheet(document, workbookPart, sheetsElement, "Sheet A", 1);

                // Save the workbook and document
                workbook.Save();
                document.Save();

                // document.Close();

                // Returns an empty array if the document isn't closed
                // stream.Position = 0; // Some places suggested this, unfortunately this doesn't seem to help.
                var bytes = stream.ToArray(); 
                File.WriteAllBytes(@"D:\sheet.xlsx", bytes);
            }
        }

        private static Sheet InsertSheet(SpreadsheetDocument doc,
                                      WorkbookPart workbookpart,
                                      Sheets sheetsElement,
                                      string name,
                                      uint id)
        {
            WorksheetPart newWorksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            string relationshipId = doc.WorkbookPart.GetIdOfPart(newWorksheetPart);

            // Give the new worksheet a name.
            if (string.IsNullOrEmpty(name))
                name = "Sheet" + id;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = id, Name = name };
            sheetsElement.AppendChild(sheet);

            return sheet;
        }
    }
}
