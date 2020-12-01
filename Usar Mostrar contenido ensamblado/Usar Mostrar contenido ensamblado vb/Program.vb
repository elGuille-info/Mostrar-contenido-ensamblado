'------------------------------------------------------------------------------
' Mostrar contenido de un ensamblado                                (01/Dic/20)
' Programa para probar la aplicación de consola 'Mostrar contenido de clases cs'
'
' (c) Guillermo (elGuille) Som, 2020
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

'Imports System
'Imports System.Data
'Imports System.Collections.Generic
'Imports System.Text
'Imports System.Linq
'Imports Microsoft.VisualBasic
'Imports vb = Microsoft.VisualBasic

Imports System

Imports gsUtilidadesNET

Module Program
    Sub Main(args As String())
        InfoEnsamblado.MostrarAyuda(True, False)

        Dim dirDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim fic = System.IO.Path.Combine(dirDocs, "Contenido gsUtilidadesNET.dll.txt")

        Dim argsDll = New String() {"E:\gsCodigo_00\Visual Studio\net core\gsUtilidadesNET\bin\Debug\net5.0-windows\gsUtilidadesNET.dll", "-c", "-m"}

        ' Si se indican parámetros de la línea de comandos
        If args.Length > 0 Then
            argsDll = args
        End If

        Dim info = InfoEnsamblado.GuardarInfo(argsDll, fic)

        If info Then
            Console.WriteLine("Todo OK.")
            Console.WriteLine($"Contenido guardado en: {fic}")
        Else
            Console.WriteLine($"Error {InfoEnsamblado.ReturnValue}")
            Console.WriteLine()

            ' Para no mostrar por la consola y tener el contenido en una variable
            'Dim infoRun = InfoEnsamblado.MostrarAyuda(False, False)
            'Console.WriteLine(infoRun)

            InfoEnsamblado.MostrarAyuda(True, False)
            Console.WriteLine("Los parámetros que se pueden usar son:")
            Console.WriteLine()

            Console.WriteLine()
        End If

        Console.WriteLine()
        Console.WriteLine("Pulsa una tecla para finalizar.")
        Console.ReadKey()
    End Sub
End Module
