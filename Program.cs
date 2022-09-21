using Aspose.Cells;
using System;

namespace NewExcel2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Workbook input = new Workbook("C:/Users/Sree Ganesh/Desktop/Excel_ To_CSV/Worldwide Rig Count Aug 2022.xlsx");

            WorksheetCollection collection = input.Worksheets;

            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {

                Worksheet worksheet = collection[worksheetIndex];

                Console.WriteLine("Worksheet: " + worksheet.Name);

                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;

                for (int i = 6; i < 35; i++)
                {

                    // Loop through each column in selected row
                    for (int j = 1; j < cols; j++)
                    {
                        
                        
                        Console.Write(worksheet.Cells[i, j].Value + " , ");// which is print the each cell data and , value.
                        //This is one more comment added from git.
                    }
                    Console.WriteLine(" ");
                }
            }
        }
    }
}
