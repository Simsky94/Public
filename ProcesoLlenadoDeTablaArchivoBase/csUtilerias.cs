using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    public class csUtilerias
    {
        public void ExceptionLog(Exception ex, String[] args)
        {
            String nomArch = "C:\\LOGS\\LOGS_EXCEP\\" + args[0].ToString() + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";

            StringBuilder logTrace = new StringBuilder();

            logTrace.AppendLine("Hora: " + DateTime.Now.ToString("hh:MM:ss"));
            for (int i = 0; i < args.Length; i++)
                logTrace.AppendLine("Parametro " + Convert.ToString(i + 1) + ": " + args[i].ToString());

            logTrace.AppendLine("Mensaje: " + ex.Message);
            logTrace.AppendLine("StackTrace: " + ex.StackTrace);
            logTrace.AppendLine("InnerException: " + ex.InnerException);
            logTrace.AppendLine("----------------------------------------------------------------------------------------------------");
            logTrace.AppendLine("----------------------------------------------------------------------------------------------------");

            EscribeLog(nomArch, logTrace.ToString());
        }

        public void EscribeLog(String Archivo, String[] args)
        {
            String nomArch = Archivo;

            StringBuilder logTrace = new StringBuilder();

            logTrace.AppendLine("Hora: " + DateTime.Now.ToString("hh:MM:ss"));
            for (int i = 0; i < args.Length; i++)
                logTrace.AppendLine("Parametro " + Convert.ToString(i + 1) + ": " + args[i].ToString());
            logTrace.AppendLine("----------------------------------------------------------------------------------------------------");
            logTrace.AppendLine("----------------------------------------------------------------------------------------------------");

            EscribeLog(nomArch, logTrace.ToString());
        }

        public void EscribeLog(String Archivo, String mensaje)
        {
            String nomArch = Archivo;

            FileInfo fi = new FileInfo(nomArch);
            if (!Directory.Exists(fi.DirectoryName))
                Directory.CreateDirectory(fi.DirectoryName);

            File.AppendAllText(nomArch, mensaje);
        }
    }
}
