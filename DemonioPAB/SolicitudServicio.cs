using System;
using System.Threading;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.IO;
using System.Data;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;
using System.Globalization;

/************************************************************************/
/* Autor: J. Antonio Barrera F.											*/
/************************************************************************/
namespace DemonioPAB
{
    public delegate void NewProcessEvent(ProcessInfo TempProcess);
    public delegate void ProcessCloseEvent(ProcessInfo TempProcess);
    public delegate void ProcessUpdateEvent(ProcessInfo TempProcess);
    
    public class SolicitudServicio
    {      
        public TcpClient Cliente = null;
        public object DatosPeticion = null;
        DVADB.DB2 dbConex;

        public SolicitudServicio(byte[] bufferRecibido, TcpClient StreamCliente, ref DVADB.DB2 dbConex)
        {
            this.dbConex = dbConex;
            if (Ejecutor.escribeLOG)
                Ejecutor.m_log.AgregaRegistro("SolicitudServicio.SolicitudServicio(byte[] bufferRecibido, TcpClient StreamCliente)");

            try
            {
                if (Ejecutor.escribeLOG)
                    Ejecutor.m_log.AgregaRegistro("SolicitudServicio.SolicitudServicio(byte[] bufferRecibido, TcpClient StreamCliente) bufferRecibido.Length: " + bufferRecibido.Length);

                //Limpia buffer
                ArrayList datos = new ArrayList();
                for (int k = 0; k < bufferRecibido.Length; k++)
                {
                    byte actual = bufferRecibido[k];
                    if (actual == Byte.MinValue)
                        break;

                    datos.Add(actual);
                }
                Byte[] myBuffer = new Byte[datos.Count];
                for (int k = 0; k < datos.Count; k++)
                {
                    myBuffer[k] = Convert.ToByte(datos[k]);
                }
                String sTemp = Encoding.ASCII.GetString(myBuffer);

                if (Ejecutor.escribeLOG)
                    Ejecutor.m_log.AgregaRegistro("SolicitudServicio.SolicitudServicio(byte[] bufferRecibido, TcpClient StreamCliente) String sTemp: " + sTemp);

                Console.WriteLine("[RECIBIDO][" + sTemp + "]");

                //try
                //{
                //DemonioPGC.Ejecutor.contProcesosAtendiendose++;

                ////EscribeMensajeArchivo(sTemp + "     CPU: " + DemonioPGC.Ejecutor.contProcesosAtendiendose);

                //if (Convert.ToDouble(DemonioPGC.Ejecutor.contProcesosAtendiendose) > 30)
                //{
                //    DVAProcesoGeneracionContabilidad.ProcesoGeneracionContabilidad enviaNuevoDemonioPGC = new DVAProcesoGeneracionContabilidad.ProcesoGeneracionContabilidad();
                //    enviaNuevoDemonioPGC.banIPAlterna = true;
                //    enviaNuevoDemonioPGC.EnviarProceso(sTemp);                    
                //    //EscribeMensajeArchivo(sTemp + "     IP ALTERNA: " + enviaNuevoDemonioPGC.strIPAlterna);                    
                //    DemonioPGC.Ejecutor.contProcesosAtendiendose--;
                //    return;
                //}
                //}
                //catch (Exception ex)
                //{   
                //    //EscribeMensajeArchivo(sTemp + "     ERROR CPU: " + ex.Message);
                //}

                DisparaProceso trigger = new DisparaProceso(sTemp);
                Thread hilo = new Thread(new ThreadStart(trigger.DispararProceso));
                hilo.Start();
            }
            catch (Exception ex)
            {
                if (Ejecutor.escribeLOG)
                {
                    Ejecutor.m_log.AgregaRegistro("Message: " + ex.Message);
                    Ejecutor.m_log.AgregaRegistro("InnerException: " + ex.InnerException);
                    Ejecutor.m_log.AgregaRegistro("StackTrace: " + ex.StackTrace);
                }
            }
        }

        //private void EscribeMensajeArchivo(String mensaje)
        //{
        //    try
        //    {
        //        String[] separador = { " " };
        //        String[] args = mensaje.Split(separador, StringSplitOptions.RemoveEmptyEntries);
        //        String nomArch = "";

        //        if(!args[0].Contains("|"))
        //            nomArch = "C:\\LOGS\\Demonio\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + args[1] + "\\" + args[0] + ".txt";
        //        else
        //            nomArch = "C:\\LOGS\\Demonio\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\" + args[1] + "\\" + args[0].Replace("|", "Comb") + ".txt";

        //        FileInfo fi = new FileInfo(nomArch);
        //        if (!Directory.Exists(fi.DirectoryName))
        //            Directory.CreateDirectory(fi.DirectoryName);

        //        StreamWriter sw = new StreamWriter(nomArch, true, Encoding.ASCII);

        //        sw.WriteLine(mensaje + "\n");

        //        sw.Close();
        //        sw.Dispose();
        //    }
        //    catch
        //    {
        //    }
        //}
    }

    public class DisparaProceso
    {   
        String mensaje = "";        

        public DisparaProceso(String sMensaje)
        {
            mensaje = sMensaje;
        }

        public void DispararProceso()
        {
            DVAConstants.Constants constantes = new DVAConstants.Constants();
            //Consultas consultas = new Consultas();

            try
            {
                if (mensaje.Trim() == "")
                    return;

                String[] separador = { " " };
                String[] args = mensaje.Split(separador, StringSplitOptions.RemoveEmptyEntries);

                //DVAControls.csLog m_log = new DVAControls.csLog("DemonioPGC", constantes.CA_GE_DIRECTORIO_LOGS, Convert.ToInt32(args[1]));

                //m_log.AgregaRegistro("[MENSAJE][" + mensaje + "]");            
                //m_log.AgregaRegistro("[ARGS][" + args.ToString() + "]");

                //ProcesoGeneracionContabilidad.Ejecutor Ejecutor = new ProcesoGeneracionContabilidad.Ejecutor();

                //ProcesoGeneracionContabilidad.Ejecutor.dtClases = DemonioPGC.Ejecutor.dtClases;
                //ProcesoGeneracionContabilidad.Ejecutor.dtCuentasMayor = DemonioPGC.Ejecutor.dtCuentasMayor;
                //ProcesoGeneracionContabilidad.Ejecutor.dtCuentasContables = DemonioPGC.Ejecutor.dtCuentasContables;
                //ProcesoGeneracionContabilidad.Ejecutor.dtConfigProceso = DemonioPGC.Ejecutor.dtConfigProceso;
                //ProcesoGeneracionContabilidad.Ejecutor.dtAtributosProceso = DemonioPGC.Ejecutor.dtAtributosProceso;
                //ProcesoGeneracionContabilidad.Ejecutor.dtProcesos = DemonioPGC.Ejecutor.dtProcesos;

                //Ejecutor.m_log = m_log;
                //m_log.AgregaRegistro("[INICIA EJECUCION]");
                ProcesoLlenadoDeTablaArchivoBase.Ejecutor Ejecutor = new ProcesoLlenadoDeTablaArchivoBase.Ejecutor();
                Ejecutor.EjecutaProceso(args);
                //m_log.AgregaRegistro("[TERMINA EJECUCION]");

                //DemonioPGC.Ejecutor.contProcesosAtendiendose--;
            }
            catch (Exception ex)
            {
                if (Ejecutor.escribeLOG)
                {
                    Ejecutor.m_log.AgregaRegistro("Message: " + ex.Message);
                    Ejecutor.m_log.AgregaRegistro("InnerException: " + ex.InnerException);
                    Ejecutor.m_log.AgregaRegistro("StackTrace: " + ex.StackTrace);
                }
            }
        }        
    }

    /*
    class CalculaCPUusage
    {
        const Process CLOSED_PROCESS = null;
        const ProcessInfo PROCESS_INFO_NOT_FOUND = null;
        
        //public event NewProcessEvent CallNewProcess;
        //public event ProcessCloseEvent CallProcessClose;
        //public event ProcessUpdateEvent CallProcessUpdate;

        public ProcessInfo[] ProcessList;
        public double CpuUsagePercent;
        public int ProcessIndex;
        public CultureInfo ValueFormat = new CultureInfo("es-MX");
        private PerformanceCounter TotalCpuUsage = new PerformanceCounter("Process", "% Processor Time", "Idle");
        private float TotalCpuUsageValue;
        public String CpuUsageDemonioPGC = "";
        public Boolean banIPAlterna = false;

        public void UpdateProcessList()
        {
            // this updates the cpu usage
            ProcessIndex = 0;
            Process[] NewProcessList = Process.GetProcesses();
            UpdateCpuUsagePercent(NewProcessList);
            AddNewProcesses(NewProcessList);
            UpdateExistingProcesses(NewProcessList);
        }

        private void UpdateCpuUsagePercent(Process[] NewProcessList)
        {
            // total the cpu usage then divide to get the usage of 1%
            double Total = 0;
            ProcessInfo TempProcessInfo;
            TotalCpuUsageValue = TotalCpuUsage.NextValue();

            foreach (Process TempProcess in NewProcessList)
            {
                //if (TempProcess.ProcessName.Contains("DemonioPGC"))
                try
                {
                    Double t = TempProcess.TotalProcessorTime.TotalMilliseconds;

                    if (TempProcess.Id == 0) continue;

                    TempProcessInfo = ProcessInfoByID(TempProcess.Id);
                    if (TempProcessInfo == PROCESS_INFO_NOT_FOUND)
                        Total += TempProcess.TotalProcessorTime.TotalMilliseconds;
                    else
                        Total += TempProcess.TotalProcessorTime.TotalMilliseconds - TempProcessInfo.OldCpuUsage;
                }
                catch (Exception ex)
                { }
            }
            CpuUsagePercent = Total / (100 - TotalCpuUsageValue);
        }

        private void AddNewProcesses(Process[] NewProcessList)
        {
            ProcessList = new ProcessInfo[NewProcessList.Length];
            // loads a new processes
            foreach (Process NewProcess in NewProcessList)
                if (!ProcessInfoExists(NewProcess))
                    AddNewProcess(NewProcess);
        }

        private void UpdateExistingProcesses(Process[] NewProcessList)
        {
            // updates the cpu usage of already loaded processes
            //if (ProcessList == null)
            //{
            //    ProcessList = new ProcessInfo[NewProcessList.Length];
            //    return;
            //}

            ProcessInfo[] TempProcessList = new ProcessInfo[NewProcessList.Length];
            ProcessIndex = 0;

            foreach (ProcessInfo TempProcess in ProcessList)
            {
                Process CurrentProcess = ProcessExists(NewProcessList, TempProcess.ID);

                //if (CurrentProcess == CLOSED_PROCESS)
                //    //CallProcessClose(TempProcess);
                //else
                //{
                TempProcessList[ProcessIndex++] = GetProcessInfo(TempProcess, CurrentProcess);
                //CallProcessUpdate(TempProcess);
                //}
            }

            ProcessList = TempProcessList;
        }

        private Process ProcessExists(Process[] NewProcessList, int ID)
        {
            // checks to see if we already loaded the process
            foreach (Process TempProcess in NewProcessList)
                if (TempProcess.Id == ID)
                    return TempProcess;

            return CLOSED_PROCESS;
        }

        private ProcessInfo GetProcessInfo(ProcessInfo TempProcess, Process CurrentProcess)
        {
            // gets the process name , id, and cpu usage
            if (CurrentProcess.Id == 0)
                TempProcess.CpuUsage = (TotalCpuUsageValue).ToString("F", ValueFormat);
            else
            {
                try
                {
                    Double t = CurrentProcess.TotalProcessorTime.TotalMilliseconds;

                    long NewCpuUsage = (long)CurrentProcess.TotalProcessorTime.TotalMilliseconds;

                    TempProcess.CpuUsage = ((NewCpuUsage - TempProcess.OldCpuUsage) / CpuUsagePercent).ToString("F", ValueFormat);
                    TempProcess.OldCpuUsage = NewCpuUsage;

                    if (CurrentProcess.ProcessName.Contains("DemonioPGC"))
                        CpuUsageDemonioPGC = TempProcess.CpuUsage;
                }
                catch (Exception ex)
                { }
            }

            return TempProcess;
        }

        private bool ProcessInfoExists(Process NewProcess)
        {
            // checks if the process info is already loaded
            if (ProcessList == null) return false;

            foreach (ProcessInfo TempProcess in ProcessList)
                if (TempProcess != PROCESS_INFO_NOT_FOUND && TempProcess.ID == NewProcess.Id)
                    return true;

            return false;
        }

        private ProcessInfo ProcessInfoByID(int ID)
        {
            // gets the process info by it's id
            if (ProcessList == null) return PROCESS_INFO_NOT_FOUND;

            for (int i = 0; i < ProcessList.Length; i++)
                if (ProcessList[i] != PROCESS_INFO_NOT_FOUND && ProcessList[i].ID == ID)
                    return ProcessList[i];

            return PROCESS_INFO_NOT_FOUND;

        }

        private void AddNewProcess(Process NewProcess)
        {
            // loads a new process
            ProcessInfo NewProcessInfo = new ProcessInfo();

            NewProcessInfo.Name = NewProcess.ProcessName;
            NewProcessInfo.ID = NewProcess.Id;

            ProcessList[ProcessIndex++] = GetProcessInfo(NewProcessInfo, NewProcess);
            //CallNewProcess(NewProcessInfo);
        }

        public void CalculaPorcentajeTiempoProcesador()
        {
            PerformanceCounter pCDemonioPGC = new PerformanceCounter();
            PerformanceCounterCategory[] perfCounters = PerformanceCounterCategory.GetCategories();
            foreach (PerformanceCounterCategory pCC in perfCounters)
            {
                if (pCC.CategoryName.Contains("Proceso"))
                {
                    string[] instanceNames;
                    System.Diagnostics.PerformanceCounterCategory mycat = new System.Diagnostics.PerformanceCounterCategory(pCC.CategoryName);
                    instanceNames = mycat.GetInstanceNames();
                    foreach (String iN in instanceNames)
                    {
                        if (iN.Contains("DemonioPGC"))
                        {
                            PerformanceCounter[] arrPC = mycat.GetCounters(iN);
                            foreach (PerformanceCounter pC in arrPC)
                            {
                                if (pC.CounterName == "% de tiempo de procesador")
                                {
                                    //for (Int32 i = 0; i <= 3; i++)
                                    //{
                                        pCDemonioPGC.CategoryName = pCC.CategoryName;
                                        pCDemonioPGC.CounterName = pC.CounterName;
                                        pCDemonioPGC.InstanceName = iN;

                                        if (banIPAlterna)
                                        {
                                            pCDemonioPGC.RawValue = pCDemonioPGC.RawValue - 5000000;
                                            CpuUsageDemonioPGC = Convert.ToDouble(pCDemonioPGC.RawValue).ToString();
                                        }
                                        else
                                            CpuUsageDemonioPGC = Convert.ToDouble(pCDemonioPGC.RawValue).ToString();

                                        //CpuUsageDemonioPGC = Math.Round(Convert.ToDouble(pCDemonioPGC.NextValue()), 2).ToString();
                                        //Thread.Sleep(1000);
                                    //}
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public class ProcessInfo
    {
        public string Name;
        public string CpuUsage;
        public int ID;
        public long OldCpuUsage;
    }
     * */
}
