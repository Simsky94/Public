using DemonioPAB;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace DemonioPAB
{
    static class Program
    {
        static void Main(string[] args)
        {
            Ejecutor ejec = new Ejecutor();
            ejec.Inicia();
        }
    }

    public class Ejecutor
    {
        public static DVAControls.csLog m_log = null;
        DVAConstants.Constants constantes = new DVAConstants.Constants();
        DVAConstants.MemoryManagement memo = new DVAConstants.MemoryManagement();
        //Consultas consultas = new Consultas();
        CalculaCPUusage calcula = new CalculaCPUusage();

        const Int32 PARAM_RUTAXML = 0;
        public static String strRutaXML = "";

        public static Boolean escribeLOG = true;

        //public static BindingList<csGECCLASETemplate> dtClases = new BindingList<csGECCLASETemplate>();
        //public static BindingList<csGECCUEMATemplate> dtCuentasMayor = new BindingList<csGECCUEMATemplate>();
        //public static BindingList<csCOCATCTSTemplate> dtCuentasContables = new BindingList<csCOCATCTSTemplate>();
        //public static BindingList<csCODCFGPRTemplate> dtConfigProceso = new BindingList<csCODCFGPRTemplate>();
        //public static BindingList<csGEDATRPRTemplate> dtAtributosProceso = new BindingList<csGEDATRPRTemplate>();
        //public static BindingList<csGECPRXMOTemplate> dtProcesos = new BindingList<csGECPRXMOTemplate>();

        public static Int32 contProcesosAtendiendose = 0;
        DVADB.DB2 dbCnx;

        public Ejecutor()
        {
            dbCnx = DVADB.DB2.Instance();
        }

        public void Inicia()
        {
            //IniciaParametros();
            IniciaLog();
            //CargaListas();
            IniciaListener();
        }

        //public void IniciaParametros()
        //{
        //    //String[] configuraciones = File.ReadAllLines(DemonioPAB.Properties.Settings.Default.RUTA_CONFIG + "\\"
        //    String[] configuraciones = File.ReadAllLines(getRutaSmartIT() + "\\"
        //            + DemonioPAB.Properties.Settings.Default.ARCHIVO_CONFIGURACION);
        //    strRutaXML = configuraciones[PARAM_RUTAXML].Substring(configuraciones[PARAM_RUTAXML].IndexOf("=") + 1);
        //}

        public void IniciaLog()
        {
            if (escribeLOG)
                m_log = new DVAControls.csLog("DemonioPAB_Inicia", constantes.CA_GE_DIRECTORIO_LOGS);
            //consultas.m_log = m_log;
            //consultas.escribeLOG = escribeLOG;
        }

        //public void CargaListas()
        //{
        //    dtClases = consultas.GetClasesRegistradas();

        //    dtCuentasMayor = consultas.GetCuentasMayor();

        //    dtConfigProceso = consultas.GetConfigProceso();

        //    dtCuentasContables = consultas.GetCuentasContables();

        //    dtAtributosProceso = consultas.GetAtributosProceso();

        //    dtProcesos = consultas.GetProcesos();
        //}

        private void IniciaListener()
        {
            ServidorAsincrono SA = new ServidorAsincrono(DemonioPAB.Properties.Settings.Default.Puerto, 45, ref dbCnx);
            //SA.m_log = m_log;
            //SA.escribeLOG = escribeLOG;

            try
            {
                SA.Iniciar();
            }
            catch (Exception ex)
            {
                if (escribeLOG)
                {
                    m_log.AgregaRegistro("Message: " + ex.Message);
                    m_log.AgregaRegistro("InnerException: " + ex.InnerException);
                    m_log.AgregaRegistro("StackTrace: " + ex.StackTrace);
                }

                try
                {
                    SA.Detener();
                }
                catch (Exception ex1)
                {
                    if (escribeLOG)
                    {
                        m_log.AgregaRegistro("Message: " + ex1.Message);
                        m_log.AgregaRegistro("InnerException: " + ex1.InnerException);
                        m_log.AgregaRegistro("StackTrace: " + ex1.StackTrace);
                    }
                }

                memo.FlushMemory();

                IniciaListener();
            }
        }

        public static String getRutaSmartIT()
        {
            String ruta = Application.ExecutablePath;
            ruta = ruta.Substring(0, ruta.LastIndexOf("\\"));
            return ruta;
        }
    }
}
