using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Globalization;

namespace DemonioPAB
{
    public class CalculaCPUusage
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
        //private PerformanceCounter TotalCpuUsage = new PerformanceCounter("Process", "% Processor Time", "Idle");
        private float TotalCpuUsageValue;
        public static Int32 CpuUsageDemonioPGC = 0;
        public static Boolean banIPAlterna = false;
        
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
            //TotalCpuUsageValue = TotalCpuUsage.NextValue();

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

                    //if (CurrentProcess.ProcessName.Contains("DemonioPGC"))
                    //    CpuUsageDemonioPGC = TempProcess.CpuUsage;
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

        public static void CalculaPorcentajeTiempoProcesador()
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

                                    //if (banIPAlterna)
                                    //{
                                    //    pCDemonioPGC.RawValue = pCDemonioPGC.RawValue - 5000000;
                                    //    CpuUsageDemonioPGC = Convert.ToDouble(pCDemonioPGC.RawValue).ToString();
                                    //}
                                    //else
                                    //    CpuUsageDemonioPGC = Convert.ToDouble(pCDemonioPGC.RawValue).ToString();

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
}
