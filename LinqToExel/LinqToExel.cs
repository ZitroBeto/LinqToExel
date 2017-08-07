using System.Collections.Generic;
using System.Linq;
using LinqToExcel;

namespace LinqToExel
{
    public class LinqToExel
    {
        private string workSheetName = "Sheet1";

        public List<Persona> ToEntidadHojaExelList(string path)
        {
            
            var book = new ExcelQueryFactory(path);
            var resultado = (from row in book.Worksheet(workSheetName)
                             let item = new Persona
                             {
                                 Id = row["Id"].Cast<string>(),
                                 Nombre = row["Nombre"].Cast<string>(),
                                 Apellido = row["Apellido"].Cast<string>()
                             }
                             select item).ToList();
            book.Dispose();
            return resultado;
        }

        public List<Persona> ToEntidadNoHeader(string path)
        {
            var book = new ExcelQueryFactory(path);
            var resultado = (from row in book.WorksheetRangeNoHeader("A6", "AD38")
                             let item = new Persona
                             {
                                 Id = row[0].Cast<string>(),
                                 Nombre = row[1].Cast<string>(),
                                 Apellido = row[2].Cast<string>()
                             }
                             select item).ToList();
            book.Dispose();
            return resultado;
        }
    }
}
