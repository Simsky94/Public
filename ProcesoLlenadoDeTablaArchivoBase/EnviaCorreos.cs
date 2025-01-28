using DVAModelsReflection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcesoLlenadoDeTablaArchivoBase
{
   
    public class EnviaCorreos
    {
        DB2Database _db = new DB2Database();
        Program escribeLog = new Program();

        public void enviarCorreo(List<string> aDestinatarios, string aAsunto, string aMensaje, List<string> aFileNames)
        {         
           try
            {
                DVAModelsReflection.SendEmail sendMail = new DVAModelsReflection.SendEmail();

                sendMail.SenderMail = _db.CA_GE_CORREO_DE_NOTIFICACIONES.ToLower().Trim();
                sendMail.Password = _db.CA_GE_CONTRASENIA_CORREO_DE_NOTIFICACIONES;
                sendMail.User = sendMail.SenderMail;

                foreach (string s in aDestinatarios)
                {
                    if (!String.IsNullOrEmpty(s))
                        sendMail.AgregaDestinatarios(s);
                }

                if (aFileNames == null || aFileNames.Count == 0)
                    sendMail.EnviaCorreoWebMail(aAsunto, aMensaje, true, System.Net.Mail.MailPriority.Normal);
                else
                    sendMail.EnviaCorreoNetMail(aAsunto, aMensaje, aFileNames, true, System.Net.Mail.MailPriority.Normal);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Program.logErr += $"-------- ERROR AL INTENTAR ENVIAR EL CORREO :   {ex.Message}"  + "\n\n\n";
                escribeLog.WriteLog(Program.logErr, 14066);
            }

        }

        private static string crearCuerpoMail(string aNombre)//mandar el id de evaluacion y id de objetivo
        {
            string temp = "<html>\r\n";

            temp += "<head>\r\n";
            temp += "<meta http-equiv=Content-Type content='text/html; charset=iso-8859-1'> \r\n";
            temp += "</head>\r\n";

            temp += "<body bgcolor='white' style='font-size: 10pt; font-family:Arial'> \r\n";

            temp += "<div style='text-align:center;'>";
            temp += "<p>Proceso Llenado reportes Archivo Base Web: " + aNombre.ToUpper() + ", Smart IT! </p>";
            temp += "</div>";

            temp += "</body>";

            temp += "</html>";

            return temp;
        }
    }

}


