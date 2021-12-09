Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Threading

Module vBoxServerMonitor

    <DllImport("kernel32")> Private Function GetPrivateProfileString(lpAppName As String, lpKeyName As String, lpDefault As String, lpReturnedString As StringBuilder, nSize As Integer, lpFileName As String) As Integer
    End Function


    ' #######################################
    ' # Definimos Variables Globales #
    ' #######################################

    Dim RUTA_EJECUCION_APLICACION As String
    Dim NOMBRE_ARCHIVO_CONFIGURACION As String
    Dim RUTA_NOMBRE_ARCHIVO_CONFIGURACION As String

    Dim SERVIDOR1_ACTIVAR As String
    Dim SERVIDOR1_IP As String
    Dim SERVIDOR1_VBOX_NAME As String
    Dim SERVIDOR1_DELAY_PING As Integer
    Dim SERVIDOR1_TIME_TO_FIRST_BOOT As Integer
    Dim SERVIDOR2_ACTIVAR As String
    Dim SERVIDOR2_IP As String
    Dim SERVIDOR2_VBOX_NAME As String
    Dim SERVIDOR2_DELAY_PING As Integer
    Dim SERVIDOR2_TIME_TO_FIRST_BOOT As Integer
    Dim SERVIDOR3_ACTIVAR As String
    Dim SERVIDOR3_IP As String
    Dim SERVIDOR3_VBOX_NAME As String
    Dim SERVIDOR3_DELAY_PING As Integer
    Dim SERVIDOR3_TIME_TO_FIRST_BOOT As Integer
    Dim GLOBAL_REINTENTOS_PING_KO As Integer
    Dim GLOBAL_VBOX_PATH As String
    Dim GLOBAL_LOAD_EXTERNAL_APP As String
    Dim GLOBAL_EXTERNAL_APP_PATH As String
    Dim GLOBAL_TIME_TO_LOAD_EXTERNAL_APP As Integer
    Dim GLOBAL_PURGUE_LOGS As Integer

    Dim VARIABLES_ARCHIVO_CONFIGURACION(21) As String

    Dim DIA_ACTUAL_1 As String
    Dim HORA_ACTUAL_1 As String
    Dim DIA_ACTUAL_2 As String
    Dim HORA_ACTUAL_2 As String
    Dim DIA_ACTUAL_3 As String
    Dim HORA_ACTUAL_3 As String

    Dim LINEA_TEXTO_NUEVA_SERVIDOR_1 As String
    Dim LINEA_TEXTO_NUEVA_SERVIDOR_2 As String
    Dim LINEA_TEXTO_NUEVA_SERVIDOR_3 As String
    Dim RUTA_ARCHIVO_DE_LOGS As String

    Dim NOMBRELOG_SERVIDOR1 As String
    Dim NOMBRELOG_SERVIDOR2 As String
    Dim NOMBRELOG_SERVIDOR3 As String



    Sub Main()

        '#################################################################
        '# Definimos variables basicas para el Funcionamiento de la App  #
        '#################################################################

        RUTA_EJECUCION_APLICACION = My.Application.Info.DirectoryPath()
        NOMBRE_ARCHIVO_CONFIGURACION = "vBoxServerMonitor.ini"
        RUTA_NOMBRE_ARCHIVO_CONFIGURACION = RUTA_EJECUCION_APLICACION & "\" & NOMBRE_ARCHIVO_CONFIGURACION

        '######################################################################
        '# Mostramos la Informacion de la App en el Arranque de la Aplicacion #
        '######################################################################

        Call MostrarInformacionApp()

        '#####################################
        '# Carga de Archivo de Configuracion #
        '#####################################

        'Leemos los parametros del Archivo de Configuracion
        VARIABLES_ARCHIVO_CONFIGURACION = FUNCION_LECTURA_ARCHIVO_CONFIGURACION(RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        'Si hay algun fallo en la Lectura del Archivo de Configuracion y nos devuelve ERROR, detenemos el programa...
        If VARIABLES_ARCHIVO_CONFIGURACION(0) = "ERROR" Then
            Exit Sub
        End If


        '###########################
        '# Inicio de la Aplicacion #
        '###########################

        ' Revisamos si existe la Carpeta y el Archivo de LOGS, y si, no, la creamos...
        RUTA_ARCHIVO_DE_LOGS = FUNCION_CREACION_DE_LA_CARPETA_LOGS(RUTA_EJECUCION_APLICACION)

        ' Generamos un nombre para el Archivo de LOG de cada Servidor
        NOMBRELOG_SERVIDOR1 = FUNCION_NOMBRE_ARCHIVO_LOG_SERVIDOR1(RUTA_EJECUCION_APLICACION)
        NOMBRELOG_SERVIDOR2 = FUNCION_NOMBRE_ARCHIVO_LOG_SERVIDOR2(RUTA_EJECUCION_APLICACION)
        NOMBRELOG_SERVIDOR3 = FUNCION_NOMBRE_ARCHIVO_LOG_SERVIDOR3(RUTA_EJECUCION_APLICACION)

        '###################
        '# Hilo Servidor 1 #
        '###################

        Dim HILO_SERVIDOR_1 As New Threading.Thread(AddressOf FUNCION_HILO_SERVIDOR_1)

        If SERVIDOR1_ACTIVAR = "SI" Then
            HILO_SERVIDOR_1 = New Threading.Thread(AddressOf FUNCION_HILO_SERVIDOR_1)
            HILO_SERVIDOR_1.Start()
        Else
            DIA_ACTUAL_1 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_1 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.DarkYellow
            Console.Write(" " & vbCrLf)
            LINEA_TEXTO_NUEVA_SERVIDOR_1 = " Monitorizacion del Servidor " & SERVIDOR1_VBOX_NAME & " deshabilitado."
            Console.Write(DIA_ACTUAL_1 & " " & HORA_ACTUAL_1 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_1 & vbCrLf)
            Console.Write(" " & vbCrLf)
        End If


        '###################
        '# Hilo Servidor 2 #
        '###################

        Dim HILO_SERVIDOR_2 As New Threading.Thread(AddressOf FUNCION_HILO_SERVIDOR_2)

        If SERVIDOR2_ACTIVAR = "SI" Then
            HILO_SERVIDOR_2 = New Threading.Thread(AddressOf FUNCION_HILO_SERVIDOR_2)
            HILO_SERVIDOR_2.Start()
        Else
            DIA_ACTUAL_2 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_2 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.DarkYellow
            Console.Write(" " & vbCrLf)
            LINEA_TEXTO_NUEVA_SERVIDOR_2 = " Monitorizacion del Servidor " & SERVIDOR2_VBOX_NAME & " deshabilitado."
            Console.Write(DIA_ACTUAL_2 & " " & HORA_ACTUAL_2 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_2 & vbCrLf)
            Console.Write(" " & vbCrLf)
        End If


        '###################
        '# Hilo Servidor 3 #
        '###################

        Dim HILO_SERVIDOR_3 As New Threading.Thread(AddressOf FUNCION_HILO_SERVIDOR_3)
        If SERVIDOR3_ACTIVAR = "SI" Then
            HILO_SERVIDOR_3 = New Threading.Thread(AddressOf FUNCION_HILO_SERVIDOR_3)
            HILO_SERVIDOR_3.Start()
        Else
            DIA_ACTUAL_3 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_3 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.DarkYellow
            Console.Write(" " & vbCrLf)
            LINEA_TEXTO_NUEVA_SERVIDOR_3 = " Monitorizacion del Servidor " & SERVIDOR3_VBOX_NAME & " deshabilitado."
            Console.Write(DIA_ACTUAL_3 & " " & HORA_ACTUAL_3 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_3 & vbCrLf)
            Console.Write(" " & vbCrLf)
        End If

        If SERVIDOR1_ACTIVAR = "NO" And SERVIDOR2_ACTIVAR = "NO" And SERVIDOR3_ACTIVAR = "NO" Then
            Threading.Thread.Sleep(10000)
        End If

    End Sub

    Sub MostrarInformacionApp()

        ' Mostramos Informacion sobre la Aplicacion
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(" " & vbCrLf)
        Console.Write("vBox Server Monitor v0.50" & vbCrLf)
        Console.Write("Code By Juan Ramon Ferrer (jr@jrferrervalero.es)" & vbCrLf)
        Console.Write("-------------------------" & vbCrLf)
        Console.Write("https://github.com/ElArtista50/" & vbCrLf)
        Console.Write("License GNU GPLv3" & vbCrLf)
        Console.Write("-------------------------" & vbCrLf)
        Console.Write("09/12/2021" & vbCrLf)
        Console.Write(" " & vbCrLf)

    End Sub

    Function FUNCION_LECTURA_ARCHIVO_CONFIGURACION(RUTA_NOMBRE_ARCHIVO_CONFIGURACION As String)

        Dim ARRAY_DATOS_DE_VUELTA(21) As String

        ' Si el Archivo de Configuracion Existe, continuamos, si no paramos...
        If System.IO.File.Exists(RUTA_NOMBRE_ARCHIVO_CONFIGURACION) = True Then
            Console.ForegroundColor = ConsoleColor.Yellow
            Console.Write("Archivo de Configuracion Localizado en:" & vbCrLf)
            Console.Write(RUTA_NOMBRE_ARCHIVO_CONFIGURACION & vbCrLf)
            Console.Write(" " & vbCrLf)
            Console.ResetColor()
        Else
            Console.ForegroundColor = ConsoleColor.Red
            Console.Write("Archivo de Configuracion INI NO encontrado." & vbCrLf)
            Console.Write("Abortando Ejecucion..." & vbCrLf)
            Console.Write(" " & vbCrLf)
            Console.ResetColor()
            Threading.Thread.Sleep(10000)
            ARRAY_DATOS_DE_VUELTA(0) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(1) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(2) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(3) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(4) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(5) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(6) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(7) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(8) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(9) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(10) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(11) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(12) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(13) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(14) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(15) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(16) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(17) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(18) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(19) = "ERROR"
            ARRAY_DATOS_DE_VUELTA(20) = "ERROR"

            GoTo SalidaFuncion
        End If

        'Leemos el archivo y volcamos la informacion en Variables...
        Dim VARIABLE_INI As New StringBuilder(64)

        ' ##############
        ' # SERVIDOR 1 #
        ' ##############

        GetPrivateProfileString("SERVIDOR1", "ACTIVAR", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR1_ACTIVAR = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR1", "IP", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR1_IP = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR1", "VBOX_NAME", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR1_VBOX_NAME = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR1", "DELAY_PING", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR1_DELAY_PING = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR1", "TIME_TO_FIRST_BOOT", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR1_TIME_TO_FIRST_BOOT = VARIABLE_INI.ToString

        ' ##############
        ' # SERVIDOR 2 #
        ' ##############

        GetPrivateProfileString("SERVIDOR2", "ACTIVAR", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR2_ACTIVAR = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR2", "IP", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR2_IP = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR2", "VBOX_NAME", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR2_VBOX_NAME = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR2", "DELAY_PING", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR2_DELAY_PING = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR2", "TIME_TO_FIRST_BOOT", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR2_TIME_TO_FIRST_BOOT = VARIABLE_INI.ToString

        ' ##############
        ' # SERVIDOR 3 #
        ' ##############

        GetPrivateProfileString("SERVIDOR3", "ACTIVAR", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR3_ACTIVAR = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR3", "IP", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR3_IP = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR3", "VBOX_NAME", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR3_VBOX_NAME = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR3", "DELAY_PING", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR3_DELAY_PING = VARIABLE_INI.ToString

        GetPrivateProfileString("SERVIDOR3", "TIME_TO_FIRST_BOOT", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        SERVIDOR3_TIME_TO_FIRST_BOOT = VARIABLE_INI.ToString

        ' ##########
        ' # GLOBAL #
        ' ##########

        GetPrivateProfileString("GLOBAL", "REINTENTOS_PING_KO", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        GLOBAL_REINTENTOS_PING_KO = VARIABLE_INI.ToString

        GetPrivateProfileString("GLOBAL", "VBOX_PATH", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        GLOBAL_VBOX_PATH = VARIABLE_INI.ToString

        GetPrivateProfileString("GLOBAL", "LOAD_EXTERNAL_APP", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        GLOBAL_LOAD_EXTERNAL_APP = VARIABLE_INI.ToString

        GetPrivateProfileString("GLOBAL", "EXTERNAL_APP_PATH", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        GLOBAL_EXTERNAL_APP_PATH = VARIABLE_INI.ToString

        GetPrivateProfileString("GLOBAL", "TIME_TO_LOAD_EXTERNAL_APP", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        GLOBAL_TIME_TO_LOAD_EXTERNAL_APP = VARIABLE_INI.ToString

        GetPrivateProfileString("GLOBAL", "PURGUE_LOGS", String.Empty, VARIABLE_INI, VARIABLE_INI.Capacity, RUTA_NOMBRE_ARCHIVO_CONFIGURACION)
        GLOBAL_PURGUE_LOGS = VARIABLE_INI.ToString

SalidaFuncion:

        Return ARRAY_DATOS_DE_VUELTA

    End Function

    Function FUNCION_CREACION_DE_LA_CARPETA_LOGS(RUTA_EJECUCION_APLICACION As String)

        Dim DATOS_DE_VUELTA As String
        Dim RUTA_CARPETA_LOGS As String

        Dim DIA_LOG As String
        DIA_LOG = DateTime.Now.ToString("yyyy-MM-dd")
        Dim HORA_LOG As String
        HORA_LOG = Now.ToString("HHmm")


        RUTA_CARPETA_LOGS = RUTA_EJECUCION_APLICACION & "\LOGS"

        ' Revisamos si la Carpeta LOGS existe ...
        If System.IO.Directory.Exists(RUTA_CARPETA_LOGS) = True Then
            'Console.ForegroundColor = ConsoleColor.Green
            'Console.Write("Carpeta LOGS Detectada, continuamos..." & vbCrLf)
            'Console.Write(" " & vbCrLf)
        Else
            Console.ForegroundColor = ConsoleColor.Red
            Console.Write("No se han detectado LOGS, creamos la carpeta en:" & vbCrLf)
            Console.Write(RUTA_CARPETA_LOGS & vbCrLf)
            Console.Write(" " & vbCrLf)
            Console.ResetColor()
            Console.Write(" " & vbCrLf)
            System.IO.Directory.CreateDirectory(RUTA_CARPETA_LOGS)
        End If

        DATOS_DE_VUELTA = RUTA_CARPETA_LOGS

        Return DATOS_DE_VUELTA

    End Function

    Function FUNCION_NOMBRE_ARCHIVO_LOG_SERVIDOR1(RUTA_EJECUCION_APLICACION As String)

        Dim DATOS_DE_VUELTA As String
        Dim RUTA_CARPETA_LOGS As String
        Dim NOMBRELOG_SERVIDOR1 As String
        Dim RUTA_CARPETA_LOGS_NOMBRELOG_SERVIDOR1 As String

        Dim DIA_LOG As String
        DIA_LOG = DateTime.Now.ToString("yyyy-MM-dd")
        Dim HORA_LOG As String
        HORA_LOG = Now.ToString("HHmm")


        RUTA_CARPETA_LOGS = RUTA_EJECUCION_APLICACION & "\LOGS"
        NOMBRELOG_SERVIDOR1 = "vBoxServerMonitor" & "_" & "SERVIDOR1" & "_" & DIA_LOG & "_" & HORA_LOG & ".txt"
        RUTA_CARPETA_LOGS_NOMBRELOG_SERVIDOR1 = RUTA_CARPETA_LOGS & "\" & NOMBRELOG_SERVIDOR1

        DATOS_DE_VUELTA = RUTA_CARPETA_LOGS_NOMBRELOG_SERVIDOR1

        Return DATOS_DE_VUELTA

    End Function

    Function FUNCION_NOMBRE_ARCHIVO_LOG_SERVIDOR2(RUTA_EJECUCION_APLICACION As String)

        Dim DATOS_DE_VUELTA As String
        Dim RUTA_CARPETA_LOGS As String
        Dim NOMBRELOG_SERVIDOR2 As String
        Dim RUTA_CARPETA_LOGS_NOMBRELOG_SERVIDOR2 As String

        Dim DIA_LOG As String
        DIA_LOG = DateTime.Now.ToString("yyyy-MM-dd")
        Dim HORA_LOG As String
        HORA_LOG = Now.ToString("HHmm")


        RUTA_CARPETA_LOGS = RUTA_EJECUCION_APLICACION & "\LOGS"
        NOMBRELOG_SERVIDOR2 = "vBoxServerMonitor" & "_" & "SERVIDOR2" & "_" & DIA_LOG & "_" & HORA_LOG & ".txt"
        RUTA_CARPETA_LOGS_NOMBRELOG_SERVIDOR2 = RUTA_CARPETA_LOGS & "\" & NOMBRELOG_SERVIDOR2

        DATOS_DE_VUELTA = RUTA_CARPETA_LOGS_NOMBRELOG_SERVIDOR2

        Return DATOS_DE_VUELTA

    End Function

    Function FUNCION_NOMBRE_ARCHIVO_LOG_SERVIDOR3(RUTA_EJECUCION_APLICACION As String)

        Dim DATOS_DE_VUELTA As String
        Dim RUTA_CARPETA_LOGS As String
        Dim NOMBRELOG_SERVIDOR3 As String
        Dim RUTA_CARPETA_LOGS_NOMBRELOG_SERVIDOR3 As String

        Dim DIA_LOG As String
        DIA_LOG = DateTime.Now.ToString("yyyy-MM-dd")
        Dim HORA_LOG As String
        HORA_LOG = Now.ToString("HHmm")


        RUTA_CARPETA_LOGS = RUTA_EJECUCION_APLICACION & "\LOGS"
        NOMBRELOG_SERVIDOR3 = "vBoxServerMonitor" & "_" & "SERVIDOR3" & "_" & DIA_LOG & "_" & HORA_LOG & ".txt"
        RUTA_CARPETA_LOGS_NOMBRELOG_SERVIDOR3 = RUTA_CARPETA_LOGS & "\" & NOMBRELOG_SERVIDOR3
        DATOS_DE_VUELTA = RUTA_CARPETA_LOGS_NOMBRELOG_SERVIDOR3

        Return DATOS_DE_VUELTA

    End Function

    Function FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA As String, NOMBRELOG As String, DIA_ACTUAL As String, HORA_ACTUAL As String)

        Dim LINEA_TEXTO_NUEVA_CON_DIA_HORA As String
        Dim ARRAY_DATOS_DE_VUELTA As String

        LINEA_TEXTO_NUEVA_CON_DIA_HORA = DIA_ACTUAL & " " & HORA_ACTUAL & " " & LINEA_TEXTO_NUEVA

        Dim LOG_WRITER As New System.IO.StreamWriter(NOMBRELOG, True)

        LOG_WRITER.WriteLine(LINEA_TEXTO_NUEVA_CON_DIA_HORA)
        LOG_WRITER.Close()

        ARRAY_DATOS_DE_VUELTA = "VACIA"

        Return ARRAY_DATOS_DE_VUELTA

    End Function

    Function FUNCION_CALCULO_DE_DIA_ACTUAL()

        Dim DIA_ACTUAL As String
        DIA_ACTUAL = DateTime.Now.ToString("yyyy-MM-dd")

        Return DIA_ACTUAL

    End Function

    Function FUNCION_CALCULO_DE_HORA_ACTUAL()

        Dim HORA_ACTUAL As String
        HORA_ACTUAL = Now.ToString("HH:mm:ss")

        Return HORA_ACTUAL

    End Function

    Function FUNCION_HILO_SERVIDOR_1()

        Dim VARIABLE_DE_VUELTA As String

        '####################################
        '# Inicio Monitorizacion Servidor 1 #
        '####################################

        ' Iniciamos visualizacion en Pantalla
        DIA_ACTUAL_1 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_1 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        LINEA_TEXTO_NUEVA_SERVIDOR_1 = "Iniciando Monitorizacion de Servidor " & SERVIDOR1_VBOX_NAME
        Console.ForegroundColor = ConsoleColor.Blue
        Console.Write(DIA_ACTUAL_1 & " " & HORA_ACTUAL_1 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_1 & vbCrLf)
        Console.Write(" " & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_1, NOMBRELOG_SERVIDOR1, DIA_ACTUAL_1, HORA_ACTUAL_1)

        DIA_ACTUAL_1 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_1 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        Console.ForegroundColor = ConsoleColor.Blue
        LINEA_TEXTO_NUEVA_SERVIDOR_1 = SERVIDOR1_VBOX_NAME & ": Mandamos Ping a la IP " & SERVIDOR1_IP
        Console.Write(DIA_ACTUAL_1 & " " & HORA_ACTUAL_1 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_1 & vbCrLf)
        Console.Write(" " & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_1, NOMBRELOG_SERVIDOR1, DIA_ACTUAL_1, HORA_ACTUAL_1)

BuclePrincipal:

        Dim CONTADOR_REINTENTOS_SERVIDOR1 As Integer

        If My.Computer.Network.Ping(SERVIDOR1_IP) = True Then
            DIA_ACTUAL_1 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_1 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Green
            LINEA_TEXTO_NUEVA_SERVIDOR_1 = "El Servidor " & SERVIDOR1_VBOX_NAME & " en la IP " & SERVIDOR1_IP & " responde al Ping."
            Console.Write(DIA_ACTUAL_1 & " " & HORA_ACTUAL_1 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_1 & vbCrLf)
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_1, NOMBRELOG_SERVIDOR1, DIA_ACTUAL_1, HORA_ACTUAL_1)
            CONTADOR_REINTENTOS_SERVIDOR1 = 0
        Else
            DIA_ACTUAL_1 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_1 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Red
            LINEA_TEXTO_NUEVA_SERVIDOR_1 = "ALERTA: El Servidor " & SERVIDOR1_VBOX_NAME & " en la IP " & SERVIDOR1_IP & " No Responde... "
            Console.Write(DIA_ACTUAL_1 & " " & HORA_ACTUAL_1 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_1 & vbCrLf)
            CONTADOR_REINTENTOS_SERVIDOR1 = CONTADOR_REINTENTOS_SERVIDOR1 + 1
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_1, NOMBRELOG_SERVIDOR1, DIA_ACTUAL_1, HORA_ACTUAL_1)
        End If

        If CONTADOR_REINTENTOS_SERVIDOR1 > 0 Then
            DIA_ACTUAL_1 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_1 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Red
            LINEA_TEXTO_NUEVA_SERVIDOR_1 = "Reintento en Servidor " & SERVIDOR1_VBOX_NAME & ": " & CONTADOR_REINTENTOS_SERVIDOR1
            Console.Write(DIA_ACTUAL_1 & " " & HORA_ACTUAL_1 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_1 & vbCrLf)
            Console.Write(" " & vbCrLf)
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_1, NOMBRELOG_SERVIDOR1, DIA_ACTUAL_1, HORA_ACTUAL_1)
        End If

        If CONTADOR_REINTENTOS_SERVIDOR1 > (GLOBAL_REINTENTOS_PING_KO - 1) Then
            DIA_ACTUAL_1 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_1 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Red
            LINEA_TEXTO_NUEVA_SERVIDOR_1 = "Numero Maximo de Reintentos Alcanzado en el Servidor " & SERVIDOR1_VBOX_NAME
            Console.Write(DIA_ACTUAL_1 & " " & HORA_ACTUAL_1 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_1 & vbCrLf)
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_1, NOMBRELOG_SERVIDOR1, DIA_ACTUAL_1, HORA_ACTUAL_1)
            GoTo ReinicioServidor
        End If

        ' Temporizamos los reintentos de PING
        Threading.Thread.Sleep((SERVIDOR1_DELAY_PING * 1000))

        GoTo BuclePrincipal

ReinicioServidor:

        DIA_ACTUAL_1 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_1 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.Write(" " & vbCrLf)
        LINEA_TEXTO_NUEVA_SERVIDOR_1 = "Procedemos a Reiniciar el Servidor " & SERVIDOR1_VBOX_NAME
        Console.Write(DIA_ACTUAL_1 & " " & HORA_ACTUAL_1 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_1 & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_1, NOMBRELOG_SERVIDOR1, DIA_ACTUAL_1, HORA_ACTUAL_1)

        ' Inicio de Maquina virtual
        Console.Write(GLOBAL_VBOX_PATH & " --startvm" & " " & SERVIDOR1_VBOX_NAME & vbCrLf)
        Console.Write(" " & vbCrLf)
        Shell(GLOBAL_VBOX_PATH & " --startvm" & " " & SERVIDOR1_VBOX_NAME)

        DIA_ACTUAL_1 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_1 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.Write(" " & vbCrLf)
        LINEA_TEXTO_NUEVA_SERVIDOR_1 = "Detenemos los Reintentos de PING en el Servidor " & SERVIDOR1_VBOX_NAME & " durante " & SERVIDOR1_TIME_TO_FIRST_BOOT & " segundos."
        Console.Write(DIA_ACTUAL_1 & " " & HORA_ACTUAL_1 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_1 & vbCrLf)
        Console.Write(" " & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_1, NOMBRELOG_SERVIDOR1, DIA_ACTUAL_1, HORA_ACTUAL_1)

        ' Detenemos los Reintentos de Ping durante el tiempo que tarda en Arrancar la Maquina Virtual
        Threading.Thread.Sleep((SERVIDOR1_TIME_TO_FIRST_BOOT * 1000))

        'Ponemos a Cero el contador de Reintentos despues de un Reinicio Forzado.
        CONTADOR_REINTENTOS_SERVIDOR1 = 0

        GoTo BuclePrincipal

        Return VARIABLE_DE_VUELTA

    End Function

    Function FUNCION_HILO_SERVIDOR_2()

        Dim VARIABLE_DE_VUELTA As String

        '####################################
        '# Inicio Monitorizacion Servidor 2 #
        '####################################

        ' Refrenamos el Hilo para que no se inicien todos a la vez y ejecute todos los comandos en plan caotico...
        Threading.Thread.Sleep(2500)


        ' Iniciamos visualizacion en Pantalla
        DIA_ACTUAL_2 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_2 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        LINEA_TEXTO_NUEVA_SERVIDOR_2 = "Iniciando Monitorizacion de Servidor " & SERVIDOR2_VBOX_NAME
        Console.ForegroundColor = ConsoleColor.Blue
        Console.Write(DIA_ACTUAL_2 & " " & HORA_ACTUAL_2 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_2 & vbCrLf)
        Console.Write(" " & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_2, NOMBRELOG_SERVIDOR2, DIA_ACTUAL_2, HORA_ACTUAL_2)

        DIA_ACTUAL_2 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_2 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        Console.ForegroundColor = ConsoleColor.Blue
        LINEA_TEXTO_NUEVA_SERVIDOR_2 = SERVIDOR2_VBOX_NAME & ": Mandamos Ping a la IP " & SERVIDOR2_IP
        Console.Write(DIA_ACTUAL_2 & " " & HORA_ACTUAL_2 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_2 & vbCrLf)
        Console.Write(" " & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_2, NOMBRELOG_SERVIDOR2, DIA_ACTUAL_2, HORA_ACTUAL_2)

BuclePrincipal:

        Dim CONTADOR_REINTENTOS_SERVIDOR2 As Integer

        If My.Computer.Network.Ping(SERVIDOR2_IP) = True Then
            DIA_ACTUAL_2 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_2 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Green
            LINEA_TEXTO_NUEVA_SERVIDOR_2 = "El Servidor " & SERVIDOR2_VBOX_NAME & " en la IP " & SERVIDOR2_IP & " responde al Ping."
            Console.Write(DIA_ACTUAL_2 & " " & HORA_ACTUAL_2 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_2 & vbCrLf)
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_2, NOMBRELOG_SERVIDOR2, DIA_ACTUAL_2, HORA_ACTUAL_2)
            CONTADOR_REINTENTOS_SERVIDOR2 = 0
        Else
            DIA_ACTUAL_2 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_2 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Red
            LINEA_TEXTO_NUEVA_SERVIDOR_2 = "ALERTA: El Servidor " & SERVIDOR2_VBOX_NAME & " en la IP " & SERVIDOR2_IP & " No Responde... "
            Console.Write(DIA_ACTUAL_2 & " " & HORA_ACTUAL_2 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_2 & vbCrLf)
            CONTADOR_REINTENTOS_SERVIDOR2 = CONTADOR_REINTENTOS_SERVIDOR2 + 1
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_2, NOMBRELOG_SERVIDOR2, DIA_ACTUAL_2, HORA_ACTUAL_2)
        End If

        If CONTADOR_REINTENTOS_SERVIDOR2 > 0 Then
            DIA_ACTUAL_2 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_2 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Red
            LINEA_TEXTO_NUEVA_SERVIDOR_2 = "Reintento en Servidor " & SERVIDOR2_VBOX_NAME & ": " & CONTADOR_REINTENTOS_SERVIDOR2
            Console.Write(DIA_ACTUAL_2 & " " & HORA_ACTUAL_2 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_2 & vbCrLf)
            Console.Write(" " & vbCrLf)
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_2, NOMBRELOG_SERVIDOR2, DIA_ACTUAL_2, HORA_ACTUAL_2)
        End If

        If CONTADOR_REINTENTOS_SERVIDOR2 > (GLOBAL_REINTENTOS_PING_KO - 1) Then
            DIA_ACTUAL_2 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_2 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Red
            LINEA_TEXTO_NUEVA_SERVIDOR_2 = "Numero Maximo de Reintentos Alcanzado en el Servidor " & SERVIDOR2_VBOX_NAME
            Console.Write(DIA_ACTUAL_2 & " " & HORA_ACTUAL_2 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_2 & vbCrLf)
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_2, NOMBRELOG_SERVIDOR2, DIA_ACTUAL_2, HORA_ACTUAL_2)
            GoTo ReinicioServidor
        End If

        ' Temporizamos los reintentos de PING
        Threading.Thread.Sleep((SERVIDOR2_DELAY_PING * 1000))

        GoTo BuclePrincipal

ReinicioServidor:

        DIA_ACTUAL_2 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_2 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.Write(" " & vbCrLf)
        LINEA_TEXTO_NUEVA_SERVIDOR_2 = "Procedemos a Reiniciar el Servidor " & SERVIDOR2_VBOX_NAME
        Console.Write(DIA_ACTUAL_2 & " " & HORA_ACTUAL_2 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_2 & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_2, NOMBRELOG_SERVIDOR2, DIA_ACTUAL_2, HORA_ACTUAL_2)

        ' Inicio de Maquina virtual
        Console.Write(GLOBAL_VBOX_PATH & " --startvm" & " " & SERVIDOR2_VBOX_NAME & vbCrLf)
        Console.Write(" " & vbCrLf)
        Shell(GLOBAL_VBOX_PATH & " --startvm" & " " & SERVIDOR2_VBOX_NAME)

        DIA_ACTUAL_2 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_2 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.Write(" " & vbCrLf)
        LINEA_TEXTO_NUEVA_SERVIDOR_2 = "Detenemos los Reintentos de PING en el Servidor " & SERVIDOR2_VBOX_NAME & " durante " & SERVIDOR2_TIME_TO_FIRST_BOOT & " segundos."
        Console.Write(DIA_ACTUAL_2 & " " & HORA_ACTUAL_2 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_2 & vbCrLf)
        Console.Write(" " & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_2, NOMBRELOG_SERVIDOR2, DIA_ACTUAL_2, HORA_ACTUAL_2)

        ' Detenemos los Reintentos de Ping durante el tiempo que tarda en Arrancar la Maquina Virtual
        Threading.Thread.Sleep((SERVIDOR2_TIME_TO_FIRST_BOOT * 1000))

        'Ponemos a Cero el contador de Reintentos despues de un Reinicio Forzado.
        CONTADOR_REINTENTOS_SERVIDOR2 = 0

        GoTo BuclePrincipal

        Return VARIABLE_DE_VUELTA

    End Function

    Function FUNCION_HILO_SERVIDOR_3()

        Dim VARIABLE_DE_VUELTA As String

        '####################################
        '# Inicio Monitorizacion Servidor 3 #
        '####################################

        ' Refrenamos el Hilo para que no se inicien todos a la vez y ejecute todos los comandos en plan caotico...
        Threading.Thread.Sleep(7500)

        ' Iniciamos visualizacion en Pantalla
        DIA_ACTUAL_3 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_3 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        LINEA_TEXTO_NUEVA_SERVIDOR_3 = "Iniciando Monitorizacion de Servidor " & SERVIDOR3_VBOX_NAME
        Console.ForegroundColor = ConsoleColor.Blue
        Console.Write(DIA_ACTUAL_3 & " " & HORA_ACTUAL_3 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_3 & vbCrLf)
        Console.Write(" " & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_3, NOMBRELOG_SERVIDOR3, DIA_ACTUAL_3, HORA_ACTUAL_3)

        DIA_ACTUAL_3 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_3 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        Console.ForegroundColor = ConsoleColor.Blue
        LINEA_TEXTO_NUEVA_SERVIDOR_3 = SERVIDOR3_VBOX_NAME & ": Mandamos Ping a la IP " & SERVIDOR3_IP
        Console.Write(DIA_ACTUAL_3 & " " & HORA_ACTUAL_3 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_3 & vbCrLf)
        Console.Write(" " & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_3, NOMBRELOG_SERVIDOR3, DIA_ACTUAL_3, HORA_ACTUAL_3)

BuclePrincipal:

        Dim CONTADOR_REINTENTOS_SERVIDOR3 As Integer

        If My.Computer.Network.Ping(SERVIDOR3_IP) = True Then
            DIA_ACTUAL_3 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_3 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Green
            LINEA_TEXTO_NUEVA_SERVIDOR_3 = "El Servidor " & SERVIDOR3_VBOX_NAME & " en la IP " & SERVIDOR3_IP & " responde al Ping."
            Console.Write(DIA_ACTUAL_3 & " " & HORA_ACTUAL_3 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_3 & vbCrLf)
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_3, NOMBRELOG_SERVIDOR3, DIA_ACTUAL_3, HORA_ACTUAL_3)
            CONTADOR_REINTENTOS_SERVIDOR3 = 0
        Else
            DIA_ACTUAL_3 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_3 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Red
            LINEA_TEXTO_NUEVA_SERVIDOR_3 = "ALERTA: El Servidor " & SERVIDOR3_VBOX_NAME & " en la IP " & SERVIDOR3_IP & " No Responde... "
            Console.Write(DIA_ACTUAL_3 & " " & HORA_ACTUAL_3 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_3 & vbCrLf)
            CONTADOR_REINTENTOS_SERVIDOR3 = CONTADOR_REINTENTOS_SERVIDOR3 + 1
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_3, NOMBRELOG_SERVIDOR3, DIA_ACTUAL_3, HORA_ACTUAL_3)
        End If

        If CONTADOR_REINTENTOS_SERVIDOR3 > 0 Then
            DIA_ACTUAL_3 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_3 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Red
            LINEA_TEXTO_NUEVA_SERVIDOR_3 = "Reintento en Servidor " & SERVIDOR3_VBOX_NAME & ": " & CONTADOR_REINTENTOS_SERVIDOR3
            Console.Write(DIA_ACTUAL_3 & " " & HORA_ACTUAL_3 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_3 & vbCrLf)
            Console.Write(" " & vbCrLf)
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_3, NOMBRELOG_SERVIDOR3, DIA_ACTUAL_3, HORA_ACTUAL_3)
        End If

        If CONTADOR_REINTENTOS_SERVIDOR3 > (GLOBAL_REINTENTOS_PING_KO - 1) Then
            DIA_ACTUAL_3 = FUNCION_CALCULO_DE_DIA_ACTUAL()
            HORA_ACTUAL_3 = FUNCION_CALCULO_DE_HORA_ACTUAL()
            Console.ForegroundColor = ConsoleColor.Red
            LINEA_TEXTO_NUEVA_SERVIDOR_3 = "Numero Maximo de Reintentos Alcanzado en el Servidor " & SERVIDOR3_VBOX_NAME
            Console.Write(DIA_ACTUAL_3 & " " & HORA_ACTUAL_3 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_3 & vbCrLf)
            FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_3, NOMBRELOG_SERVIDOR3, DIA_ACTUAL_3, HORA_ACTUAL_3)
            GoTo ReinicioServidor
        End If

        ' Temporizamos los reintentos de PING
        Threading.Thread.Sleep((SERVIDOR3_DELAY_PING * 1000))

        GoTo BuclePrincipal

ReinicioServidor:

        DIA_ACTUAL_3 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_3 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.Write(" " & vbCrLf)
        LINEA_TEXTO_NUEVA_SERVIDOR_3 = "Procedemos a Reiniciar el Servidor " & SERVIDOR3_VBOX_NAME
        Console.Write(DIA_ACTUAL_3 & " " & HORA_ACTUAL_3 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_3 & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_3, NOMBRELOG_SERVIDOR3, DIA_ACTUAL_3, HORA_ACTUAL_3)

        ' Inicio de Maquina virtual
        Console.Write(GLOBAL_VBOX_PATH & " --startvm" & " " & SERVIDOR3_VBOX_NAME & vbCrLf)
        Console.Write(" " & vbCrLf)
        Shell(GLOBAL_VBOX_PATH & " --startvm" & " " & SERVIDOR3_VBOX_NAME)

        DIA_ACTUAL_3 = FUNCION_CALCULO_DE_DIA_ACTUAL()
        HORA_ACTUAL_3 = FUNCION_CALCULO_DE_HORA_ACTUAL()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.Write(" " & vbCrLf)
        LINEA_TEXTO_NUEVA_SERVIDOR_3 = "Detenemos los Reintentos de PING en el Servidor " & SERVIDOR3_VBOX_NAME & " durante " & SERVIDOR3_TIME_TO_FIRST_BOOT & " segundos."
        Console.Write(DIA_ACTUAL_3 & " " & HORA_ACTUAL_3 & " " & LINEA_TEXTO_NUEVA_SERVIDOR_3 & vbCrLf)
        Console.Write(" " & vbCrLf)
        FUNCION_ESCRIBIMOS_EN_EL_LOG(LINEA_TEXTO_NUEVA_SERVIDOR_3, NOMBRELOG_SERVIDOR3, DIA_ACTUAL_3, HORA_ACTUAL_3)

        ' Detenemos los Reintentos de Ping durante el tiempo que tarda en Arrancar la Maquina Virtual
        Threading.Thread.Sleep((SERVIDOR3_TIME_TO_FIRST_BOOT * 1000))

        'Ponemos a Cero el contador de Reintentos despues de un Reinicio Forzado.
        CONTADOR_REINTENTOS_SERVIDOR3 = 0

        GoTo BuclePrincipal

        Return VARIABLE_DE_VUELTA

    End Function

End Module
