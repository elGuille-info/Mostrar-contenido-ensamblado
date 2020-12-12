'-----------------------------------------------------------------------------
' Mostrar contenido de clases usando Reflection versión de VB      (01/Dic/20)
'
' Convertido de C# a VB con CSharpToVB de Paul1956
'   No pone Option Infer On y en los bucles no indica el tipo de datos
'   Main lo ha declarado como Private
'
' (c) Guillermo (elGuille) Som, 2020
'-----------------------------------------------------------------------------

Option Compare Text
Option Explicit On
Option Infer On
Option Strict On

Imports System.Text
Imports System.Linq
Imports System.Collections.Generic
Imports System.Reflection
Imports System.IO
Imports System.Diagnostics

#If ESX86 Then
namespace gsUtilidadesNETx86
#Else
Namespace gsUtilidadesNET
#End If

    Public Class InfoEnsamblado
        Shared mostrarTodo As Boolean
        Shared mostrarClases As Boolean
        Shared mostrarPropiedades As Boolean
        Shared verbose As Boolean
        Shared mostrarMetodos As Boolean

        Public Shared Property ReturnValue As Integer

        Shared Function Main(args As String()) As Integer
            Console.WriteLine("Mostrar el contenido de una clase usando Reflection. [Versión para VB]")
            Console.WriteLine()

            Console.WriteLine(InfoTipo(args), False)

            Console.WriteLine("Pulsa una tecla para terminar.")
            Console.ReadKey()

            Return ReturnValue
        End Function

        ''' <summary>
        ''' Guardar la información del ensamblado indicado en los argumentos
        ''' (los mismos que si se ejecuta desde la línea de comandos).
        ''' </summary>
        ''' <param name="args">Los argumentos a usar para informar del ensamblado.</param>
        ''' <param name="fic">El nombre del fichero donde se guardará la información.</param>
        ''' <returns>Devuelve True si todo fue bien, false si hubo errores.</returns>
        ''' <remarks>El tipo de error se averigua con InfoEnsamblado.ReturnValue</remarks>
        ''' <remarks>El formato usado para guardar es Latin1</remarks>
        Public Shared Function GuardarInfo(args As String(), fic As String) As Boolean
            Dim res As String = InfoTipo(args)
            If ReturnValue > 0 Then
                Return False
            End If
            Using sw As System.IO.StreamWriter = New IO.StreamWriter(fic, False, Encoding.Latin1)
                sw.WriteLine(res)
                sw.Flush()
                sw.Close()
            End Using

            Return True
        End Function

        ''' <summary>
        ''' Devuelve la información del ensamblado indicado.
        ''' Se aceptan los mísmos argumentos que en la llamada desde la línea de comandos.
        ''' </summary>
        ''' <param name="args">Los argumentos a procesar</param>
        ''' <param name="mostrarComandos">Si se muestran los argumentos usados.</param>
        ''' <returns>Una cadena con la iformación o los errores producidos.</returns>
        Public Shared Function InfoTipo(args As String(), Optional mostrarComandos As Boolean = False) As String
            Dim sb As System.Text.StringBuilder = New StringBuilder

            If args.Length = 0 Then
                ReturnValue = 1
                Return MostrarAyuda(True, False)
            End If

            Dim nombreEnsamblado As String = args(0)
            If String.IsNullOrEmpty(nombreEnsamblado) Then
                ReturnValue = 2
                Return MostrarAyuda(True, False)
            End If

            If mostrarComandos Then
                sb.AppendLine("Línea de comandos:")
                'System.IO.Path.GetFileName(nombreEnsamblado)
                sb.Append("    ")
                sb.Append(IO.Path.GetFileName(nombreEnsamblado))
                sb.Append(" ")
                For i = 1 To args.Length - 1
                    sb.Append(args(i))
                    sb.Append("")
                Next
                'sb.AppendLine(string.Join(' ', args));
                sb.AppendLine()
                sb.AppendLine()
            End If

            Dim tipo As String = ""

            mostrarTodo = True
            ' Estos dos valores solo se tienen en cuenta si mostrarTodo es false
            mostrarClases = False
            mostrarPropiedades = False
            mostrarMetodos = False
            verbose = True

            Dim opChars As Char() = New Char() {"-"c, "/"c}

            For i = 1 To args.Length - 1
                Dim op As String = args(i).ToLower()

                Dim opSinopChars As String = op.TrimStart(opChars)
                If opSinopChars.StartsWith("h") OrElse opSinopChars.StartsWith("?") OrElse opSinopChars.StartsWith("help") Then
                    ReturnValue = 0
                    Return MostrarAyuda(True, False)
                End If
                ' Si se indica input
                If opSinopChars.StartsWith("i") OrElse opSinopChars.StartsWith("input") Then
                    Console.WriteLine("Indica el path al ensamblado (.dll) a examinar: ")
                    nombreEnsamblado = Console.ReadLine()
                    If Not File.Exists(nombreEnsamblado) Then
                        ReturnValue = 2
                        Return "El ensamblado indicado no se encuentra."
                    End If
                End If
                ' si se indica tipo
                If opSinopChars.StartsWith("t") OrElse opSinopChars.StartsWith("tipo") Then
                    Dim j As Integer = args(i).IndexOf(":")
                    If j = -1 Then
                        ReturnValue = 5
                        Return MostrarAyuda(True, False)
                    End If
                    tipo = args(i).Substring(j + 1)
                    If String.IsNullOrEmpty(tipo) Then
                        ReturnValue = 5
                        Return MostrarAyuda(True, False)
                    End If
                End If
                If opSinopChars.StartsWith("v") OrElse opSinopChars.StartsWith("verbose") Then
                    verbose = True
                End If

                If opSinopChars.StartsWith("p") OrElse opSinopChars.StartsWith("property") Then
                    mostrarPropiedades = True
                    'mostrarMetodos = false;
                    mostrarTodo = False
                    'mostrarClases = false;
                End If
                If opSinopChars.StartsWith("m") OrElse opSinopChars.StartsWith("method") Then
                    mostrarMetodos = True
                    'mostrarPropiedades = false;
                    mostrarTodo = False
                    'mostrarClases = false;
                End If
                If opSinopChars.StartsWith("pm") Then
                    mostrarPropiedades = True
                    mostrarMetodos = True
                    mostrarTodo = False
                    'mostrarClases = false;
                End If
                If opSinopChars.StartsWith("c") OrElse opSinopChars.StartsWith("class") Then
                    mostrarClases = True
                    'mostrarPropiedades = false;
                    mostrarTodo = False
                End If
            Next

            ' Carga el ensamblado y mostrar el contenido pedido
            Dim objAssembly As Assembly
            Try
                objAssembly = Assembly.LoadFrom(nombreEnsamblado)
                If objAssembly Is Nothing Then
                    ReturnValue = 3
                    Return "Error al cargar el ensamblado."
                End If
            Catch ex As Exception
                ReturnValue = -1
                Return ex.Message
            End Try

            sb.AppendLine($"Contenido del ensamblado '{IO.Path.GetFileName(nombreEnsamblado)}'")
            sb.AppendLine()

            Dim losTipos As Type()

            ' esto da error al cargar los ensamblados de Windows.Forms
            ' o tipos que no están definidos
            ' losTipos = objAssembly.GetTypes();

            ' Ejemplo tomado de:
            ' https://haacked.com/archive/2012/07/23/get-all-types-in-an-assembly.aspx/
            Try
                losTipos = objAssembly.GetTypes()
            Catch ex As ReflectionTypeLoadException
                losTipos = ex.Types
            End Try

            Dim elTipo As Type = Nothing
            If tipo.Any() Then
                elTipo = objAssembly.[GetType](tipo)
            End If

            Dim indent As Integer = 0

            If Not (elTipo Is Nothing) Then
                ' (losTipos is null && !(elTipo is null))
                Dim t As System.Type = elTipo
                sb.AppendLine("Información del tipo indicado.")
                sb.AppendLine()
                MostrarInfoTipo(sb, indent, t)
            Else
                sb.AppendLine("Información de los tipos definidos en el ensamblado.")
                sb.AppendLine()
                For i = 0 To losTipos.Count() - 1
                    Dim t As System.Type = losTipos(i)
                    If Not (t Is Nothing) Then
                        indent = MostrarInfoTipo(sb, indent, t)
                        sb.AppendLine()
                    End If
                Next
            End If
            ReturnValue = 0
            Return sb.ToString().TrimEnd()
        End Function

        Private Shared Function MostrarInfoTipo(sb As StringBuilder, indent As Integer, t As Type) As Integer
            If mostrarClases OrElse mostrarTodo Then
                If t.IsEnum Then
                    sb.AppendLine($"Enumeración: {t.Name}")
                ElseIf t.IsInterface Then
                    sb.AppendLine($"Interface: {t.Name}")
                ElseIf t.IsClass Then
                    sb.AppendLine($"Clase: {t.Name}")
                ElseIf t.IsValueType Then
                    sb.AppendLine($"ValueType: {t.Name}")
                End If

                If t.IsEnum Then
                    Dim enumNames As String() = t.GetEnumNames()
                    If enumNames.Length > 0 Then
                        indent += 4
                        For i = 0 To enumNames.Length - 1
                            sb.AppendLine($"{" ".PadLeft(indent)}{enumNames(i)}")
                        Next
                        indent -= 4
                        sb.AppendLine()
                    End If
                    ' Produce los mismos resultados que GetEnumNames
                    'var enumV = t.GetEnumValues();
                    'if (enumV.Length > 0)
                    '{
                    '    indent += 4;
                    '    for (var i = 0; i < enumV.Length; i++)
                    '    {
                    '        sb.AppendLine($"{" ".PadLeft(indent)}{enumV.GetValue(i)}");
                    '    }
                    '    indent -= 4;
                    '    sb.AppendLine();
                    '}
                End If

                If verbose Then
                    If t.IsGenericType Then
                        indent += 4
                        sb.AppendLine($"{" ".PadLeft(indent)}IsGenericType = {t.IsGenericType}")
                        indent -= 4
                    End If

                    ' Los constructores
                    Dim constrInfo As ConstructorInfo() = t.GetConstructors()
                    If constrInfo.Length > 0 Then
                        indent += 4
                        sb.AppendLine($"{" ".PadLeft(indent)}Constructores:")
                        indent += 4
                        For j = 0 To constrInfo.Length - 1
                            If constrInfo(j).IsPrivate Then
                                sb.AppendLine($"{" ".PadLeft(indent)}IsPrivate = {constrInfo(j).IsPrivate}")
                            End If
                            If constrInfo(j).IsPublic Then
                                sb.AppendLine($"{" ".PadLeft(indent)}IsPublic = {constrInfo(j).IsPublic}")
                            End If
                            If constrInfo(j).IsAbstract Then
                                sb.AppendLine($"{" ".PadLeft(indent)}IsAbstract = {constrInfo(j).IsAbstract}")
                            End If
                            If constrInfo(j).IsStatic Then
                                sb.AppendLine($"{" ".PadLeft(indent)}IsStatic = {constrInfo(j).IsStatic}")
                            End If
                            If constrInfo(j).IsVirtual Then
                                sb.AppendLine($"{" ".PadLeft(indent)}IsVirtual = {constrInfo(j).IsVirtual}")
                            End If

                            Dim parInfo As System.Reflection.ParameterInfo() = constrInfo(j).GetParameters()
                            If parInfo.Length > 0 Then
                                indent += 4
                                sb.Append($"{" ".PadLeft(indent)}Parámetros: ")
                                For k = 0 To parInfo.Length - 1
                                    If k > 0 Then
                                        sb.Append(", ")
                                    End If
                                    If parInfo(k).IsOptional Then
                                        sb.Append($"{parInfo(k).IsOptional}")
                                    End If
                                    If parInfo(k).IsOut Then
                                        sb.Append($"{parInfo(k).IsOut}")
                                    End If
                                    If parInfo(k).IsRetval Then
                                        sb.Append($"{parInfo(k).IsRetval}")
                                    End If
                                    sb.Append($"{parInfo(k).ParameterType.Name.Replace("System.", "")} ")
                                    sb.Append($"{parInfo(k).Name}")
                                Next
                                sb.AppendLine()
                                indent -= 4
                            Else sb.AppendLine($"{" ".PadLeft(indent + 4)}Sin parámetros")
                            End If
                        Next
                        indent -= 8
                    End If
                End If
            End If
            ' Los campos 
            Dim campos As System.Reflection.FieldInfo() = t.GetFields()
            If Not t.IsEnum AndAlso campos.Length > 0 AndAlso mostrarTodo Then
                indent += 4
                sb.AppendLine($"{" ".PadLeft(indent)}{t.Name}.Campos:")
                indent += 4
                For j = 0 To campos.Length - 1
                    sb.Append($"{" ".PadLeft(indent)}{campos(j).FieldType.Name.Replace("System.", "")}")
                    sb.AppendLine($" {campos(j).Name}")
                    If verbose Then
                        If campos(j).IsPrivate Then
                            sb.AppendLine($"{" ".PadLeft(indent)}IsPrivate = {campos(j).IsPrivate}")
                        End If
                        If campos(j).IsPublic Then
                            sb.AppendLine($"{" ".PadLeft(indent)}IsPublic = {campos(j).IsPublic}")
                        End If
                        If campos(j).IsStatic Then
                            sb.AppendLine($"{" ".PadLeft(indent)}IsStatic = {campos(j).IsStatic}")
                        End If
                        If campos(j).IsInitOnly Then
                            sb.AppendLine($"{" ".PadLeft(indent)}IsInitOnly = {campos(j).IsInitOnly}")
                        End If
                    End If
                Next
                indent -= 8
            End If
            ' Las propiedades
            Dim propiedades As System.Reflection.PropertyInfo() = t.GetProperties()
            If propiedades.Length > 0 AndAlso (mostrarPropiedades OrElse mostrarTodo) Then
                indent += 4
                sb.AppendLine($"{" ".PadLeft(indent)}{t.Name}.Propiedades:")
                indent += 4
                For j = 0 To propiedades.Length - 1
                    sb.Append($"{" ".PadLeft(indent)}{propiedades(j).PropertyType.Name.Replace("System.", "")}")
                    sb.AppendLine($" {propiedades(j).Name}")
                    'sb.AppendLine($"{" ".PadLeft(indent)}{propiedades[j].Name}");

                    If verbose Then
                        indent += 4
                        If propiedades(j).CanRead Then
                            sb.AppendLine($"{" ".PadLeft(indent)}CanRead: {propiedades(j).CanRead}")
                        End If
                        If propiedades(j).CanWrite Then
                            sb.AppendLine($"{" ".PadLeft(indent)}CanWrite: {propiedades(j).CanWrite}")
                        End If
                        indent -= 4
                    End If
                Next
                indent -= 8
            End If
            ' Los métodos
            Dim metodos As System.Reflection.MethodInfo() = t.GetMethods()
            If metodos.Length > 0 AndAlso (mostrarMetodos OrElse mostrarTodo) Then
                indent += 4
                sb.AppendLine($"{" ".PadLeft(indent)}{t.Name}.Métodos:")
                indent += 4
                For j = 0 To metodos.Length - 1
                    'if (metodos[j].IsHideBySig) break;
                    If metodos(j).Name.StartsWith("get_") OrElse metodos(j).Name.StartsWith("set_") Then
                        Continue For
                    End If

                    'sb.Append($"{" ".PadLeft(indent)}{metodos[j].MemberType}");
                    Try
                        If String.IsNullOrEmpty(metodos(j).Name) Then
                            Continue For
                        End If
                        ' El tipo del método
                        sb.Append($"{" ".PadLeft(indent)}{metodos(j).ReturnParameter.ToString().Replace("System.", "")}")
                        sb.AppendLine($" {metodos(j).Name}")
                        'sb.AppendLine($"{" ".PadLeft(indent)}{metodos[j].Name}");
                    Catch

                    End Try

                    If verbose Then
                        ' Mostrar los argumentos
                        Try
                            Dim parInfo As System.Reflection.ParameterInfo() = metodos(j).GetParameters()

                            If parInfo.Length > 0 Then
                                indent += 4
                                sb.Append($"{" ".PadLeft(indent)}Parámetros: ")
                                For k = 0 To parInfo.Length - 1
                                    If k > 0 Then
                                        sb.Append(", ")
                                    End If
                                    If parInfo(k).IsOptional Then
                                        sb.Append($"[Optional] ")
                                    End If
                                    If parInfo(k).IsOut Then
                                        sb.Append($"[Out] ")
                                    End If
                                    If parInfo(k).IsRetval Then
                                        sb.Append($"[Retval] ")
                                    End If
                                    sb.Append($"{parInfo(k).ParameterType.Name.Replace("System.", "")} ")
                                    sb.Append($"{parInfo(k).Name}")
                                    If parInfo(k).IsOptional Then
                                        sb.Append($" = {parInfo(k).DefaultValue}")
                                    End If
                                Next
                                sb.AppendLine()
                                indent -= 4
                            End If
                        Catch

                        End Try

                    End If
                Next
                indent -= 8
            End If
            ' Las interfaces
            Dim interfaces As System.Type() = t.GetInterfaces()
            If interfaces.Length > 0 Then
                indent += 4
                sb.AppendLine($"{" ".PadLeft(indent)}{t.Name}.Interfaces:")
                indent += 4
                For j = 0 To interfaces.Length - 1
                    ' El tipo de
                    'sb.Append($"{" ".PadLeft(indent)}{interfaces[j]. .ReturnParameter.ToString().Replace("System.", "")}");
                    'sb.AppendLine($" {interfaces[j].Name}");
                    sb.AppendLine($"{" ".PadLeft(indent)}{interfaces(j).Name}")
                    Dim miembros As System.Reflection.MemberInfo() = interfaces(j).GetMembers()
                    If miembros.Length > 0 Then
                        indent += 4
                        For k = 0 To miembros.Length - 1
                            sb.AppendLine($"{" ".PadLeft(indent)}{miembros(k).Name}")
                        Next
                        indent -= 4
                    End If
                Next
                indent -= 8
            End If

            Return indent
        End Function

        ''' <summary>
        ''' Mostrar la ayuda de este programa.
        ''' </summary>
        ''' <param name="esperar"></param>
        ''' <returns></returns>
        Public Shared Function MostrarAyuda(mostrarEnConsola As Boolean, esperar As Boolean) As String
            Dim ayudaMsg As String = $"{VersionInfo()}
Opciones de la lína de comandos:{ProductName} 
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
   -1   Otro error no definido."

            If mostrarEnConsola Then
                Console.WriteLine(ayudaMsg)
            End If
            If esperar Then
                Console.ReadKey()
            End If

            Return ayudaMsg
        End Function

        Shared ProductName As String
        Shared ProductVersion As String
        Shared FileVersion As String

        Private Shared Function VersionInfo() As String
            Dim ensamblado As System.Reflection.Assembly = Assembly.GetExecutingAssembly()
            Dim fvi As System.Diagnostics.FileVersionInfo = FileVersionInfo.GetVersionInfo(ensamblado.Location)

            ProductName = fvi.ProductName
            ProductVersion = fvi.ProductVersion
            FileVersion = fvi.FileVersion

            Return $"{ProductName} v{ProductVersion} ({FileVersion})"
        End Function
    End Class
End Namespace
