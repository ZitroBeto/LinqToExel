using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LinqToExel
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Persona> lst = new List<Persona>();
            LinqToExel lqe = new LinqToExel();
            //string path = @"C:\Users\roberto.ortiz\Documents\Work In Progress\Softtek\Visual Studio 2015\Projects\LinqToExel\LinqToExel\test.xlsx";
            string path = @"C:\Users\roberto.ortiz\Documents\Work In Progress\Softtek\Visual Studio 2015\Projects\LinqToExel\LinqToExel\testNoHeaders.xlsx";
            lst = lqe.ToEntidadNoHeader(path);//ToEntidadHojaExelList(path);
            foreach(Persona p in lst)
            {
                Console.Write("Id:{0}, Nombre:{1}, Apellido:{2}", p.Id,p.Nombre,p.Apellido);
                Console.WriteLine();
            }
            Console.ReadLine();

        }
    }
}
