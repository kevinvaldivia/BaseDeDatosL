using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Collections.Generic;


//Como ejecutar el programa:
//1er paso: Primero debe verificar y cambiar la ruta del archivo en la LINEA 17: EL MISMO REPO se encuentra el archivo Excel (Boca.xlsx)
//2do paso: Despues de configuar la ruta, debe abrir una terminal e ir al archivo BaseDeDatosL: cd BaseDeDatosL
//3er paso: En la terminal, se debe añadir el package de DataOleDb: dotnet add package System.Data.OleDb
//4to paso: En la terminal, se debe compilar el proyecto con el comando: dotnet buil
//5to paso: En la terminal, se debe ejecutar el proyecto con el comando: dotnet run

class Program
{
    static void Main(string[] args)
    {
        //ruta del archivo Excel

        string archivoExcel = @"C:\Users\kevin\OneDrive\Escritorio\Boca.xlsx";

        //cadena de conexion
        string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={archivoExcel};Extended Properties='Excel 12.0 XML';";

        //Lista de los archivos excel
        List<Registro> registros = new List<Registro>();

        //establezco conexion y carga datos del excel
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            //abrimos conexion
            connection.Open();

            //consulta para seleccionar todos los datos de la hoja
            string consulta = "SELECT * FROM [Hoja1$]";

            //creamos el commando con la consulta y conexion
            using (OleDbCommand command = new OleDbCommand(consulta, connection))
            {
                //ejecutamos la consulta y el reader
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    //leemos los datos y los agregamos a la lista
                    while (reader.Read()) 
                    {
                        registros.Add(new Registro
                            {
                            Nombre = reader["Nombre"] != DBNull.Value ? reader["Nombre"].ToString() : null,
                            Apellido = reader["Apellido"] != DBNull.Value ? reader["Apellido"].ToString() : null,
                            Dorsal = reader["Dorsal"] != DBNull.Value ? Convert.ToInt32(reader["Dorsal"]) : 0,
                        });
                    }
                }
            }
            foreach (var registro in registros)
            {
                Console.WriteLine($"Nombre: {registro.Nombre}, Apellido: {registro.Apellido}, Dorsal: {registro.Dorsal}");
            }
            //Usando LINQ
            //filtrando y seleccionando un dato con un input 
            Console.Write("Escribir el jugador que desea ver, la primera letra debe estar en Mayuscula: ");
            string buscarJugador = Console.ReadLine();
            var jugadoresEncontrado = registros.Where(r => r.Nombre == buscarJugador).ToList();
            foreach(var jugador in jugadoresEncontrado)
            {
                Console.Write("Jugador encontrado: ");
                Console.Write($"Nombre:{jugador.Nombre}, Apellido:{jugador.Apellido}, Dorsal: {jugador.Dorsal}");
            }
                //USANDO LINQ para hacer una consulta todos los jugadores con el dorsal menores e igual 5
            var jugadores = registros.Where(r => r.Dorsal <= 5);
            foreach (var jugador in jugadores)
            {
                Console.WriteLine("JUGADORES CON DORSALES MENORES E IGUAL A 5");
                Console.WriteLine($"Nombre:{jugador.Nombre}, Apellido:{jugador.Apellido}, Dorsal: {jugador.Dorsal}");
            }
        }
        
        
    }

    //definino clase para representar cada registro del archivo excel
    public class Registro
    {
        public string Nombre { get; set; }
        public string Apellido { get; set; }
        public int Dorsal { get; set; }
    }
}
