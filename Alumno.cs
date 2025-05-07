using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _6_de_Mayo
{
    internal class Alumno
    {
        public Alumno(string nombre, int edad, string carrera, int matricula, DateTime fechanacimiento)
        {
            Nombre = nombre;
            Edad = edad;
            Carrera = carrera;
            Matricula = matricula;
            Fechanacimiento = fechanacimiento;
        }

        public String Nombre { get; set; }
        public int Edad { get; set; }

        public string Carrera { get; set; }

        public int Matricula { get; set; }

        public DateTime Fechanacimiento { get; set; }
    }
}
