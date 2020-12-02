//-----------------------------------------------------------------------------
// Mostrar contenido de clases usando Reflection versión de C#      (01/Dic/20)
//
//
// (c) Guillermo (elGuille) Som, 2020
//-----------------------------------------------------------------------------

using System;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Diagnostics;

#if ESX86
namespace gsUtilidadesNETx86
#else
namespace gsUtilidadesNET
#endif
{
    public class InfoEnsamblado
    {
        static bool mostrarTodo;
        static bool mostrarClases;
        static bool mostrarPropiedades;
        static bool verbose;
        static bool mostrarMetodos;
        //private static int returnValue;
        public static int ReturnValue { get; set; }

        static int Main(string[] args)
        {
            Console.WriteLine("Mostrar el contenido de una clase usando Reflection. [Versión para C#]");
            Console.WriteLine();

            Console.WriteLine(InfoTipo(args), false);
            
            Console.WriteLine("Pulsa una tecla para terminar.");
            Console.ReadKey();

            return ReturnValue;
        }

        /// <summary>
        /// Guardar la información del ensamblado indicado en los argumentos
        /// (los mismos que si se ejecuta desde la línea de comandos).
        /// </summary>
        /// <param name="args">Los argumentos a usar para informar del ensamblado.</param>
        /// <param name="fic">El nombre del fichero donde se guardará la información.</param>
        /// <returns>Devuelve True si todo fue bien, false si hubo errores.</returns>
        /// <remarks>El tipo de error se averigua con InfoEnsamblado.ReturnValue</remarks>
        /// <remarks>El formato usado para guardar es Latin1</remarks>
        public static bool GuardarInfo(string[] args, string fic)
        {
            var res = InfoTipo(args);
            if (ReturnValue > 0)
                return false;

            using (var sw = new System.IO.StreamWriter(fic, false, Encoding.Latin1))
            {
                sw.WriteLine(res);
                sw.Flush();
                sw.Close();
            }
            return true;
        }

        /// <summary>
        /// Devuelve la información del ensamblado indicado.
        /// Se aceptan los mísmos argumentos que en la llamada desde la línea de comandos.
        /// </summary>
        /// <param name="args">Los argumentos a procesar</param>
        /// <param name="mostrarComandos">Si se muestran los argumentos usados.</param>
        /// <returns>Una cadena con la iformación o los errores producidos.</returns>
        public static string InfoTipo(string[] args, bool mostrarComandos = false)
        {
            var sb = new StringBuilder();

            if (args.Length == 0)
            {
                ReturnValue = 1;
                return MostrarAyuda(true, false);
            }

            string nombreEnsamblado = args[0];
            if (string.IsNullOrEmpty(nombreEnsamblado))
            {
                ReturnValue = 2;
                return MostrarAyuda(true, false);
            }

            if (mostrarComandos)
            {
                sb.AppendLine("Línea de comandos:");
                //System.IO.Path.GetFileName(nombreEnsamblado)
                sb.Append("    ");
                sb.Append(System.IO.Path.GetFileName(nombreEnsamblado));
                sb.Append(" ");
                for (var i = 1; i < args.Length; i++)
                {
                    sb.Append(args[i]);
                    sb.Append("");
                }
                //sb.AppendLine(string.Join(' ', args));
                sb.AppendLine();
                sb.AppendLine();
            }

            string tipo = "";

            mostrarTodo = true;
            // Estos dos valores solo se tienen en cuenta si mostrarTodo es false
            mostrarClases = false;
            mostrarPropiedades = false;
            mostrarMetodos = false;
            verbose = true;

            var opChars = new char[] { '-', '/' };

            for (var i = 1; i < args.Length; i++)
            {
                string op = args[i].ToLower();

                var opSinopChars = op.TrimStart(opChars);
                if (opSinopChars.StartsWith("h") || opSinopChars.StartsWith("?") || opSinopChars.StartsWith("help"))
                {
                    ReturnValue = 0;
                    return MostrarAyuda(true, false);
                }
                // Si se indica input
                if (opSinopChars.StartsWith("i") || opSinopChars.StartsWith("input"))
                {
                    Console.WriteLine("Indica el path al ensamblado (.dll) a examinar: ");
                    nombreEnsamblado = Console.ReadLine();
                    if (!File.Exists(nombreEnsamblado))
                    {
                        ReturnValue = 2;
                        return "El ensamblado indicado no se encuentra.";
                    }
                }
                // si se indica tipo
                if (opSinopChars.StartsWith("t") || opSinopChars.StartsWith("tipo"))
                {
                    var j = args[i].IndexOf(":");
                    if (j == -1)
                    {
                        ReturnValue = 5;
                        return MostrarAyuda(true, false);
                    }
                    tipo = args[i].Substring(j + 1);
                    if (string.IsNullOrEmpty(tipo))
                    {
                        ReturnValue = 5;
                        return MostrarAyuda(true, false);
                    }
                }
                if (opSinopChars.StartsWith("v") || opSinopChars.StartsWith("verbose"))
                    verbose = true;

                if (opSinopChars.StartsWith("p") || opSinopChars.StartsWith("property"))
                {
                    mostrarPropiedades = true;
                    //mostrarMetodos = false;
                    mostrarTodo = false;
                    //mostrarClases = false;
                }
                if (opSinopChars.StartsWith("m") || opSinopChars.StartsWith("method"))
                {
                    mostrarMetodos = true;
                    //mostrarPropiedades = false;
                    mostrarTodo = false;
                    //mostrarClases = false;
                }
                if (opSinopChars.StartsWith("pm"))
                {
                    mostrarPropiedades = true;
                    mostrarMetodos = true;
                    mostrarTodo = false;
                    //mostrarClases = false;
                }
                if (opSinopChars.StartsWith("c") || opSinopChars.StartsWith("class"))
                {
                    mostrarClases = true;
                    //mostrarPropiedades = false;
                    mostrarTodo = false;
                }
            }

            // Carga el ensamblado y mostrar el contenido pedido
            Assembly objAssembly;
            try
            {
                objAssembly = Assembly.LoadFrom(nombreEnsamblado);
                if (objAssembly == null)
                {
                    ReturnValue = 3;
                    return "Error al cargar el ensamblado.";
                }
            }
            catch (Exception ex)
            {
                ReturnValue = -1;
                return ex.Message;
            };

            sb.AppendLine($"Contenido del ensamblado '{System.IO.Path.GetFileName(nombreEnsamblado)}'");
            sb.AppendLine();

            Type[] losTipos;

            // esto da error al cargar los ensamblados de Windows.Forms
            // o tipos que no están definidos
            // losTipos = objAssembly.GetTypes();

            // Ejemplo tomado de:
            // https://haacked.com/archive/2012/07/23/get-all-types-in-an-assembly.aspx/
            try
            {
                losTipos = objAssembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                losTipos = ex.Types;
            }

            Type elTipo = null;
            if (tipo.Any())
                elTipo = objAssembly.GetType(tipo);

            var indent = 0;

            if (!(elTipo is null)) // (losTipos is null && !(elTipo is null))
            {
                var t = elTipo;
                sb.AppendLine("Información del tipo indicado.");
                sb.AppendLine();
                MostrarInfoTipo(sb, indent, t);
            }
            else
            {
                sb.AppendLine("Información de los tipos definidos en el ensamblado.");
                sb.AppendLine();
                for (var i = 0; i < losTipos.Count(); i++)
                {
                    var t = losTipos[i];
                    if (!(t is null)) 
                    {
                        indent = MostrarInfoTipo(sb, indent, t);
                        sb.AppendLine();
                    }
                        
                }
            }
            ReturnValue = 0;
            return sb.ToString().TrimEnd();
        }

        static int MostrarInfoTipo(StringBuilder sb, int indent, Type t)
        {
            if (mostrarClases || mostrarTodo)
            {
                if (t.IsEnum)
                    sb.AppendLine($"Enumeración: {t.Name}");
                else if(t.IsInterface)
                    sb.AppendLine($"Interface: {t.Name}");
                else if(t.IsClass)
                    sb.AppendLine($"Clase: {t.Name}");
                else if(t.IsValueType)
                    sb.AppendLine($"ValueType: {t.Name}");

                if (t.IsEnum)
                {
                    var enumNames = t.GetEnumNames();
                    if (enumNames.Length > 0)
                    {
                        indent += 4;
                        for (var i = 0; i < enumNames.Length; i++)
                        {
                            sb.AppendLine($"{" ".PadLeft(indent)}{enumNames[i]}");
                        }
                        indent -= 4;
                        sb.AppendLine();
                    }
                    // Produce los mismos resultados que GetEnumNames
                    //var enumV = t.GetEnumValues();
                    //if (enumV.Length > 0)
                    //{
                    //    indent += 4;
                    //    for (var i = 0; i < enumV.Length; i++)
                    //    {
                    //        sb.AppendLine($"{" ".PadLeft(indent)}{enumV.GetValue(i)}");
                    //    }
                    //    indent -= 4;
                    //    sb.AppendLine();
                    //}
                }

                if (verbose)
                {
                    if (t.IsGenericType)
                    {
                        indent += 4;
                        sb.AppendLine($"{" ".PadLeft(indent)}IsGenericType = {t.IsGenericType}");
                        indent -= 4;
                    }

                    // Los constructores
                    ConstructorInfo[] constrInfo = t.GetConstructors();
                    if (constrInfo.Length > 0)
                    {
                        indent += 4;
                        sb.AppendLine($"{" ".PadLeft(indent)}Constructores:");
                        indent += 4;
                        for (var j = 0; j < constrInfo.Length; j++)
                        {
                            if (constrInfo[j].IsPrivate)
                                sb.AppendLine($"{" ".PadLeft(indent)}IsPrivate = {constrInfo[j].IsPrivate}");
                            if (constrInfo[j].IsPublic)
                                sb.AppendLine($"{" ".PadLeft(indent)}IsPublic = {constrInfo[j].IsPublic}");
                            if (constrInfo[j].IsAbstract)
                                sb.AppendLine($"{" ".PadLeft(indent)}IsAbstract = {constrInfo[j].IsAbstract}");
                            if (constrInfo[j].IsStatic)
                                sb.AppendLine($"{" ".PadLeft(indent)}IsStatic = {constrInfo[j].IsStatic}");
                            if (constrInfo[j].IsVirtual)
                                sb.AppendLine($"{" ".PadLeft(indent)}IsVirtual = {constrInfo[j].IsVirtual}");

                            var parInfo = constrInfo[j].GetParameters();
                            if (parInfo.Length > 0)
                            {
                                indent += 4;
                                sb.Append($"{" ".PadLeft(indent)}Parámetros: ");
                                for (var k=0; k < parInfo.Length; k++)
                                {
                                    if (k > 0)
                                        sb.Append(", ");
                                    if(parInfo[k].IsOptional)
                                        sb.Append($"{parInfo[k].IsOptional}");
                                    if(parInfo[k].IsOut)
                                        sb.Append($"{parInfo[k].IsOut}");
                                    if(parInfo[k].IsRetval)
                                        sb.Append($"{parInfo[k].IsRetval}");
                                    sb.Append($"{parInfo[k].ParameterType.Name.Replace("System.","")} ");
                                    sb.Append($"{parInfo[k].Name}");
                                }
                                sb.AppendLine();
                                indent -= 4;
                            }
                            else
                                sb.AppendLine($"{" ".PadLeft(indent+4)}Sin parámetros");
                        }
                        indent -= 8;
                    }
                }
            }
            // Los campos 
            var campos = t.GetFields();
            if (!t.IsEnum && campos.Length > 0 && mostrarTodo)
            {
                indent += 4;
                sb.AppendLine($"{" ".PadLeft(indent)}{t.Name}.Campos:");
                indent += 4;
                for (var j = 0; j < campos.Length; j++)
                {
                    sb.Append($"{" ".PadLeft(indent)}{campos[j].FieldType.Name.Replace("System.", "")}");
                    sb.AppendLine($" {campos[j].Name}");
                    if (verbose)
                    {
                        if (campos[j].IsPrivate)
                            sb.AppendLine($"{" ".PadLeft(indent)}IsPrivate = {campos[j].IsPrivate}");
                        if (campos[j].IsPublic)
                            sb.AppendLine($"{" ".PadLeft(indent)}IsPublic = {campos[j].IsPublic}");
                        if (campos[j].IsStatic)
                            sb.AppendLine($"{" ".PadLeft(indent)}IsStatic = {campos[j].IsStatic}");
                        if (campos[j].IsInitOnly)
                            sb.AppendLine($"{" ".PadLeft(indent)}IsInitOnly = {campos[j].IsInitOnly}");
                    }
                }
                indent -= 8;
            }
            // Las propiedades
            var propiedades = t.GetProperties();
            if (propiedades.Length > 0 && (mostrarPropiedades || mostrarTodo))
            {
                indent += 4;
                sb.AppendLine($"{" ".PadLeft(indent)}{t.Name}.Propiedades:");
                indent += 4;
                for (var j = 0; j < propiedades.Length; j++)
                {
                    sb.Append($"{" ".PadLeft(indent)}{propiedades[j].PropertyType.Name.Replace("System.", "")}");
                    sb.AppendLine($" {propiedades[j].Name}");
                    //sb.AppendLine($"{" ".PadLeft(indent)}{propiedades[j].Name}");

                    if (verbose)
                    {
                        indent += 4;
                        if (propiedades[j].CanRead)
                            sb.AppendLine($"{" ".PadLeft(indent)}CanRead: {propiedades[j].CanRead}");
                        if (propiedades[j].CanWrite)
                            sb.AppendLine($"{" ".PadLeft(indent)}CanWrite: {propiedades[j].CanWrite}");
                        indent -= 4;
                    }
                }
                indent -= 8;
            }
            // Los métodos
            var metodos = t.GetMethods();
            if (metodos.Length > 0 && (mostrarMetodos || mostrarTodo))
            {
                indent += 4;
                sb.AppendLine($"{" ".PadLeft(indent)}{t.Name}.Métodos:");
                indent += 4;
                for (var j = 0; j < metodos.Length; j++)
                {
                    //if (metodos[j].IsHideBySig) break;
                    if (metodos[j].Name.StartsWith("get_") || metodos[j].Name.StartsWith("set_"))
                        continue;

                    //sb.Append($"{" ".PadLeft(indent)}{metodos[j].MemberType}");
                    try
                    {
                        if (string.IsNullOrEmpty(metodos[j].Name))
                            continue;
                        // El tipo del método
                        sb.Append($"{" ".PadLeft(indent)}{metodos[j].ReturnParameter.ToString().Replace("System.", "")}");
                        sb.AppendLine($" {metodos[j].Name}");
                        //sb.AppendLine($"{" ".PadLeft(indent)}{metodos[j].Name}");
                    }
                    catch { };

                    if (verbose)
                    {
                        // Mostrar los argumentos
                        try
                        {
                            var parInfo = metodos[j].GetParameters();

                            if (parInfo.Length > 0)
                            {
                                indent += 4;
                                sb.Append($"{" ".PadLeft(indent)}Parámetros: ");
                                for (var k = 0; k < parInfo.Length; k++)
                                {
                                    if (k > 0)
                                        sb.Append(", ");
                                    if (parInfo[k].IsOptional)
                                        sb.Append($"[Optional] ");
                                    if (parInfo[k].IsOut)
                                        sb.Append($"[Out] ");
                                    if (parInfo[k].IsRetval)
                                        sb.Append($"[Retval] ");
                                    sb.Append($"{parInfo[k].ParameterType.Name.Replace("System.", "")} ");
                                    sb.Append($"{parInfo[k].Name}");
                                    if (parInfo[k].IsOptional)
                                        sb.Append($" = {parInfo[k].DefaultValue}");
                                }
                                sb.AppendLine();
                                indent -= 4;
                            }
                        }
                        catch { };
                    }
                }
                indent -= 8;
            }
            // Las interfaces
            var interfaces = t.GetInterfaces();
            if (interfaces.Length > 0)
            {
                indent += 4;
                sb.AppendLine($"{" ".PadLeft(indent)}{t.Name}.Interfaces:");
                indent += 4;
                for (var j = 0; j < interfaces.Length; j++)
                {
                    // El tipo de
                    //sb.Append($"{" ".PadLeft(indent)}{interfaces[j]. .ReturnParameter.ToString().Replace("System.", "")}");
                    //sb.AppendLine($" {interfaces[j].Name}");
                    sb.AppendLine($"{" ".PadLeft(indent)}{interfaces[j].Name}");
                    var miembros = interfaces[j].GetMembers();
                    if (miembros.Length > 0)
                    {
                        indent += 4;
                        for (var k = 0; k < miembros.Length; k++)
                        {
                            sb.AppendLine($"{" ".PadLeft(indent)}{miembros[k].Name}");
                        }
                        indent -= 4;
                    }
                }
                indent -= 8;
            }

            return indent;
        }

        /// <summary>
        /// Mostrar la ayuda de este programa.
        /// </summary>
        /// <param name="esperar"></param>
        /// <returns></returns>
        public static string MostrarAyuda(bool mostrarEnConsola, bool esperar)
        {
            var ayudaMsg = @$"{VersionInfo()}
Opciones de la lína de comandos:
{ProductName} 
    ensamblado [opciones]

ensamblado
    El path del ensamblado a analizar.
    Si el path contiene espacios hay que encerrarlo entre comillas dobles.

opciones            Las opciones se pueden indicar con - o /
    h ? help        Muestra esta ayuda
    [t]ipo:nombre   Indicar el tipo a mostrar (no hay separación entre : y el nombre).
                    El tipo (incluido con el espacio de nombres) del que se mostrará la información
                        Por ejemplo: /t:gsUtilidadesNET.Marcadores
    a[ll]           Muestra todo el contenido del ensamblado. 
                        Predeterminado = si. 
    c[lass]         Muestra solo las clases.
                        Predeterminado = no.
    p[roperty]      Muestra solo las propiedades.
                        Predeterminado = no.
    m[ethod]        Muestra solo los métodos.
                        Predeterminado = no.
    pm              Muestra las propiedades y métodos.
                        Predeterminado = no.
                    Nota:
                    Las opciones -c -p -m o -pm se pueden combinar para mostrar los tipos que queramos.

    v[erbose]       Muestra detalles de las clases y propiedades/métodos:
                    En las clases muestra los constructores
                    En los métodos/propiedades muestra los argumentos.
                        Predeterminado = si
    i[nput]         Preguntar por el nombre del ensamblado a usar.

Valores devueltos:
    0   Todo fue bien.
    1   No se han indicado parámetros en la línea de comandos.   
    2   No se encuentra el ensamblado indicado.
        O no se ha indicado en la línea de comandos como primer argumento.
    3   Error al cargar o procesar el ensamblado.
    5   No se ha indicado el tipo.
   -1   Otro error no definido.
";
            
            if(mostrarEnConsola)
                Console.WriteLine(ayudaMsg);
            if (esperar)
                Console.ReadKey();
            
            return ayudaMsg;
        }

        static string ProductName;
        static string ProductVersion;
        static string FileVersion;

        static string VersionInfo()
        {
            var ensamblado = Assembly.GetExecutingAssembly();
            var fvi = FileVersionInfo.GetVersionInfo(ensamblado.Location);

            ProductName = fvi.ProductName;
            ProductVersion = fvi.ProductVersion;
            FileVersion = fvi.FileVersion;

            return $"{ProductName} v{ProductVersion} ({FileVersion})";
        }
    }
}
