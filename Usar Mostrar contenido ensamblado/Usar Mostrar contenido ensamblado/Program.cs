//-----------------------------------------------------------------------------
// Mostrar contenido de un ensamblado                               (01/Dic/20)
// Programa para probar la aplicación de consola 'Mostrar contenido de clases cs'
//
// (c) Guillermo (elGuille) Som, 2020
//-----------------------------------------------------------------------------

using System;

using gsUtilidadesNET;

namespace Usar_Mostrar_contenido_ensamblado
{
    class Program
    {
        static void Main(string[] args)
        {
            _ = InfoEnsamblado.MostrarAyuda(true, false);

            var dirDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var fic = System.IO.Path.Combine(dirDocs, "Contenido gsUtilidadesNET.dll.txt");

            var argsDll = new string[] { @"E:\gsCodigo_00\Visual Studio\net core\gsUtilidadesNET\bin\Debug\net5.0-windows\gsUtilidadesNET.dll", "-c", "-m" };
            
            // Si se indican parámetros de la línea de comandos
            if (args.Length > 0)
                argsDll = args;

            var info = InfoEnsamblado.GuardarInfo(argsDll, fic);
            if (info)
            {
                Console.WriteLine("Todo OK.");
                Console.WriteLine($"Contenido guardado en: {fic}");
            }
            else
            {
                Console.WriteLine($"Error {InfoEnsamblado.ReturnValue}");
                Console.WriteLine();
                Console.WriteLine("Los parámetros que se pueden usar son:");
                Console.WriteLine();

                // Para no mostrar por la consola y tener el contenido en una variable
                //var infoRun = InfoEnsamblado.MostrarAyuda(false, false);
                //Console.WriteLine(infoRun);

                _ = InfoEnsamblado.MostrarAyuda(true, false);
                Console.WriteLine();
            }
            Console.WriteLine();
            Console.WriteLine("Pulsa una tecla para finalizar.");
            Console.ReadKey();
        }
    }
}
