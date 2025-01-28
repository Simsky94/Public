using System;
using System.Collections.Generic;
using DVAModelsReflection;
using DVAModelsReflectionFINA.Models.FINA;
using DVAModelsReflection.Models.CONT;
using DVAModelsReflection.Models.GRAL;
using System.Linq;
using DVAModelsReflection.Models.AUSA;
using DVAModelsReflection.Models.NOM;
using System.Diagnostics;
using System.Threading;
using DVAModelsReflection.Models.SEGU;
using System.IO;
using System.Text;
using System.ComponentModel;
using System.Windows.Forms;
using System.Net.NetworkInformation;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    public class Program
    {
        public static string RutaServidor = "";
        public static string RutaExcelArchivoBaseCorreo = "";
        public static string logErr = "";
        public static string logFolderPath = "";
        public static string logFileName = "";
        public static string logFilePath = ""; 
        public static string nombreArchivo = "";
        public static int UsuarioLog = 0;
        public static DVAControls.csLog m_log = null;
        public static Boolean escribeLOG = true;

        public static void Main(string[] args)
        {
            try
            {
                DB2Database _db = new DB2Database();
                #region Log
                Program escribeLog = new Program();
                var getRutaLogs = ParametrosProcesosFina.BuscarPorNombre(_db, "LogsRutaArchivosBase");
                if (getRutaLogs != null)
                {
                    logFolderPath = getRutaLogs.ValorParametro;
                }
                else
                {
                  logFolderPath = @"\\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase\LogProcesoConsola\";
                }

                var getRutaServidor = ParametrosProcesosFina.BuscarPorNombre(_db, "ArchivosBaseRuta");
                if (getRutaServidor != null)
                {
                    RutaServidor = getRutaServidor.ValorParametro;
                    //RutaServidor = @"C:\Sistemas DVA\Finanzas\ProcesoLlenadoDeTablaArchivoBase\ProcesoLlenadoDeTablaArchivoBase\bin\Debug\ProcesoLlenadoDeTablaArchivoBase.exe";
                }
                else
                {
                    RutaServidor = @"\\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase\";
       
                    //"C:\Sistemas DVA\Finanzas\ProcesoLlenadoDeTablaArchivoBase\ProcesoLlenadoDeTablaArchivoBase\bin\Debug\ProcesoLlenadoDeTablaArchivoBase.exe"
                }

                string nombreBase = "Log_";
                string msgErr = "";
                #endregion

                if (args[0] == "INCADEA") //INCADEA 2022 12 300
                {
                    ProcesoDeLlenadoInfoINCADEALayOut procesoINCADEA = new ProcesoDeLlenadoInfoINCADEALayOut(_db,
                        Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), Convert.ToInt32(args[3]));
                    procesoINCADEA.LecturaExcel();
                }
                else if (args[0] == "ARCHIVO_BASE_EXCEL") //ARCHIVO_BASE_EXCEL 2024 12 RM V1 O ARCHIVO_BASE_EXCEL 2024 12 RM V1 28
                {
                    #region ARCHIVO_BASE_EXCEL

                    if (args.Length == 6)
                    {
                        int idAgencia = Convert.ToInt32(args[5]);
                        Agencia agencia = Agencia.Buscar(_db, idAgencia);

                        if (agencia.Id == 588)
                            agencia.Siglas = "AFAC";
                        else if (agencia.Id == 596)
                            agencia.Siglas = "DUC";
                        else if (agencia.Id == 286)
                            agencia.Siglas = "LYM";
                        else if (agencia.Id == 212)
                            agencia.Siglas = "EAH";
                        else if (agencia.Id == 348)
                            agencia.Siglas = "EAP";
                        else if (agencia.Id == 255)
                            agencia.Siglas = "EAZ";
                        else if (agencia.Id == 646)
                            agencia.Siglas = "LAR";
                        else if (agencia.Id == 585)
                            agencia.Siglas = "LMP";
                        else if (agencia.Id == 300)
                            agencia.Siglas = "MMB";
                        else if (agencia.Id == 365)
                            agencia.Siglas = "MMC";
                        else if (agencia.Id == 366)
                            agencia.Siglas = "MMR";
                        else if (agencia.Id == 373)
                            agencia.Siglas = "MMAB";

                        ProcesoDeLecturaArchivoBaseExcelYLlenadoInfo proceso = new ProcesoDeLecturaArchivoBaseExcelYLlenadoInfo(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, args[4].ToString(), agencia.Siglas, args[3].ToString());
                        proceso.LecturaExcel();
                    }
                    else
                    {
                        List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1);
                        List<int> aIdAgencias = new List<int>();
                        aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                        List<Agencia> agencias = Agencia.ListarPorIds(_db, aIdAgencias);

                        foreach (Agencia agencia in agencias)
                        {
                            if (agencia.Id == 588)
                                agencia.Siglas = "AFAC";
                            else if (agencia.Id == 596)
                                agencia.Siglas = "DUC";
                            else if (agencia.Id == 286)
                                agencia.Siglas = "LYM";
                            else if (agencia.Id == 212)
                                agencia.Siglas = "EAH";
                            else if (agencia.Id == 348)
                                agencia.Siglas = "EAP";
                            else if (agencia.Id == 255)
                                agencia.Siglas = "EAZ";
                            else if (agencia.Id == 646)
                                agencia.Siglas = "LAR";
                            else if (agencia.Id == 585)
                                agencia.Siglas = "LMP";
                            else if (agencia.Id == 300)
                                agencia.Siglas = "MMB";
                            else if (agencia.Id == 365)
                                agencia.Siglas = "MMC";
                            else if (agencia.Id == 366)
                                agencia.Siglas = "MMR";
                            else if (agencia.Id == 373)
                                agencia.Siglas = "MMAB";

                            ProcesoDeLecturaArchivoBaseExcelYLlenadoInfo proceso = new ProcesoDeLecturaArchivoBaseExcelYLlenadoInfo(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, args[4].ToString(), agencia.Siglas, args[3].ToString());
                            proceso.LecturaExcel();
                        }
                    }

                    #endregion
                }
                else if (args[0] == "ARCHIVO_BASE_EXCEL_VS_WEB") //ARCHIVO_BASE_EXCEL_VS_WEB 2024 11 RM O ARCHIVO_BASE_EXCEL_VS_WEB 2024 11 RM 28
                {
                    #region ARCHIVO_BASE_EXCEL_VS_WEB

                    //string[] ids = args[3].ToString().Split(sep);
                    //List<int> ids = args[3].ToString().Split(sep).ToListInt();
                    ProcesoGeneraComparacionExcelVsWeb proceso;

                    if (args.Length == 5)
                    {
                        char sep = ',';
                        List<int> idsAgencia = args[4].ToString().Split(sep).ToListInt();

                        proceso = new ProcesoGeneraComparacionExcelVsWeb(_db, Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), args[3].ToString(), idsAgencia); // todas las agencias
                        if (args[3].ToString() == "RM")
                            proceso.GeneraExcelRM(idsAgencia);
                        else if (args[3].ToString() == "BG")
                            proceso.GeneraExcelBG(idsAgencia);
                        else if (args[3].ToString() == "VP")
                            proceso.GeneraExcelVP(idsAgencia);
                        else if (args[3].ToString() == "Cuentas")
                            proceso.GeneraExcelCuentas(idsAgencia, true);
                        else if (args[3].ToString() == "AG")
                            proceso.GeneraExcelAG(idsAgencia, true);
                        else if (args[3].ToString() == "SC")
                            proceso.GeneraExcelSC(idsAgencia, true);
                    }
                    else
                    {
                        proceso = new ProcesoGeneraComparacionExcelVsWeb(_db, Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), args[3].ToString()); //por agencia
                        if (args[3].ToString() == "RM")
                            proceso.GeneraExcelRM(null);
                        else if (args[3].ToString() == "BG")
                            proceso.GeneraExcelBG(null);
                        else if (args[3].ToString() == "VP")
                            proceso.GeneraExcelVP(null);
                        else if (args[3].ToString() == "Cuentas")
                            proceso.GeneraExcelCuentas(null, false);
                        else if (args[3].ToString() == "AG")
                            proceso.GeneraExcelAG(null, false);
                        else if (args[3].ToString() == "SC")
                            proceso.GeneraExcelSC(null, false);
                    }

                    #endregion
                }
                else if (args[0] == "ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_ANTERIOR") //ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_ANTERIOR 2024 11 V1 O ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_ANTERIOR 2024 11 V1 28 O ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_ANTERIOR 2024 11 RM V1 28
                {
                    #region ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_ANTERIOR

                    List<string> pestanias = new List<string> { "BG", "OPL", "PX", "RMAN", "RM", "Cuentas", "VP", "AG", "CC", "AF", "SI", "SC",
                    "GANMT", "GANMF", "GANMVI", "GGNM", "GASM", "GSM", "GHPM", "GRM", "GDM" }; //"D"

                    if (args.Length == 6)
                    {
                        int idAgencia = Convert.ToInt32(args[5]);
                        Agencia agencia = Agencia.Buscar(_db, idAgencia);

                        if (agencia.Id == 588)
                            agencia.Siglas = "AFAC";
                        else if (agencia.Id == 596)
                            agencia.Siglas = "DUC";
                        else if (agencia.Id == 286)
                            agencia.Siglas = "LYM";
                        else if (agencia.Id == 212)
                            agencia.Siglas = "EAH";
                        else if (agencia.Id == 348)
                            agencia.Siglas = "EAP";
                        else if (agencia.Id == 255)
                            agencia.Siglas = "EAZ";
                        else if (agencia.Id == 646)
                            agencia.Siglas = "LAR";
                        else if (agencia.Id == 585)
                            agencia.Siglas = "LMP";
                        else if (agencia.Id == 300)
                            agencia.Siglas = "MMB";
                        else if (agencia.Id == 365)
                            agencia.Siglas = "MMC";
                        else if (agencia.Id == 366)
                            agencia.Siglas = "MMR";
                        else if (agencia.Id == 373)
                            agencia.Siglas = "MMAB";

                        nombreArchivo = $"{nombreBase}_0_{agencia.Id}__{DateTime.Now:yy-MM-dd}.txt";

                        ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[4].ToString(), args[3].ToString(), agencia.Nombre);
                        proceso.LlenaExcel();
                    }
                    else if (args.Length == 5)
                    {
                        int idAgencia = Convert.ToInt32(args[4]);
                        Agencia agencia = Agencia.Buscar(_db, idAgencia);

                        if (agencia.Id == 588)
                            agencia.Siglas = "AFAC";
                        else if (agencia.Id == 596)
                            agencia.Siglas = "DUC";
                        else if (agencia.Id == 286)
                            agencia.Siglas = "LYM";
                        else if (agencia.Id == 212)
                            agencia.Siglas = "EAH";
                        else if (agencia.Id == 348)
                            agencia.Siglas = "EAP";
                        else if (agencia.Id == 255)
                            agencia.Siglas = "EAZ";
                        else if (agencia.Id == 646)
                            agencia.Siglas = "LAR";
                        else if (agencia.Id == 585)
                            agencia.Siglas = "LMP";
                        else if (agencia.Id == 300)
                            agencia.Siglas = "MMB";
                        else if (agencia.Id == 365)
                            agencia.Siglas = "MMC";
                        else if (agencia.Id == 366)
                            agencia.Siglas = "MMR";
                        else if (agencia.Id == 373)
                            agencia.Siglas = "MMAB";

                        nombreArchivo = $"{nombreBase}_0_{agencia.Id}__{DateTime.Now:yy-MM-dd}.txt";

                        ProcesoDeLlenadoExcelDesdeWeb proceso;

                        foreach (string pestania in pestanias)
                        {
                            proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[3].ToString(), pestania, agencia.Nombre);
                            proceso.LlenaExcel();
                        }

                        pestanias = new List<string> { "GANMT", "GANMF", "GANMVI", "GGNM", "GASM", "GSM", "GHPM", "GRM", "GDM", "RM" }; //"D"

                        foreach (string pestania in pestanias)
                        {
                            proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[3].ToString(), pestania, agencia.Nombre);
                            proceso.LlenaExcelAjustaDiferencia();
                        }
                    }
                    else
                    {
                        List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1);
                        List<int> aIdAgencias = new List<int>();
                        aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                        List<Agencia> agencias = Agencia.ListarPorIds(_db, aIdAgencias);

                        foreach (Agencia agencia in agencias)
                        {
                            if (agencia.Id == 588)
                                agencia.Siglas = "AFAC";
                            else if (agencia.Id == 596)
                                agencia.Siglas = "DUC";
                            else if (agencia.Id == 286)
                                agencia.Siglas = "LYM";
                            else if (agencia.Id == 212)
                                agencia.Siglas = "EAH";
                            else if (agencia.Id == 348)
                                agencia.Siglas = "EAP";
                            else if (agencia.Id == 255)
                                agencia.Siglas = "EAZ";
                            else if (agencia.Id == 646)
                                agencia.Siglas = "LAR";
                            else if (agencia.Id == 585)
                                agencia.Siglas = "LMP";
                            else if (agencia.Id == 300)
                                agencia.Siglas = "MMB";
                            else if (agencia.Id == 365)
                                agencia.Siglas = "MMC";
                            else if (agencia.Id == 366)
                                agencia.Siglas = "MMR";
                            else if (agencia.Id == 373)
                                agencia.Siglas = "MMAB";

                            nombreArchivo = $"{nombreBase}_0_{agencia.Id}__{DateTime.Now:yy-MM-dd}.txt";

                            ProcesoDeLlenadoExcelDesdeWeb proceso;

                            foreach (string pestania in pestanias)
                            {
                                proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                    Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[3].ToString(), pestania, agencia.Nombre);
                                proceso.LlenaExcel();
                            }

                            pestanias = new List<string> { "GANMT", "GANMF", "GANMVI", "GGNM", "GASM", "GSM", "GHPM", "GRM", "GDM", "RM" }; //"D"

                            foreach (string pestania in pestanias)
                            {
                                proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                    Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[3].ToString(), pestania, agencia.Nombre);
                                proceso.LlenaExcelAjustaDiferencia();
                            }
                        }
                    }

                    #endregion
                }
                else if (args[0] == "ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB") //ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB 2024 11 V1 O ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB 2024 11 V1 28 O ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB 2024 11 RM V1 28
                {
                    #region ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB

                    List<EnvioCorreosProcesoLlenadoArchivoBaseWeb> consultaInfoProceso = new List<EnvioCorreosProcesoLlenadoArchivoBaseWeb>();
                    //ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB 2024 5 RM V1 28
                    int __numPrograma = 2468;
                    int numReportes = 0;
                    var idUsuario = 0;
                    var __anio = 0;
                    var __mes = 0;
                    var __siglasPestania = "";
                    var __version = "";
                    var __idAgencia = 0;
                    var __claveUnica = 0;
                    var __siglasAgencia = "";
                    var __agenciaNombre = "";
                    bool __existRegProceso = false;
                    bool __esTodos = false;
                    bool __TerminaProcesoTodos = false;


                    if (args.Length != 5)
                    {
                        string[] idAgenciaYusuario = args[5].Split('_');
                        numReportes = 0;
                        idUsuario = 0;
                        __anio = Convert.ToInt32(args[1]);
                        __mes = Convert.ToInt32(args[2]);
                        __siglasPestania = args[3];
                        __version = args[4];
                        __idAgencia = Convert.ToInt32(idAgenciaYusuario[0]);

                        nombreArchivo = $"{nombreBase}_0_{__idAgencia}__{DateTime.Now:yy-MM-dd}.txt";

                        consultaInfoProceso = EnvioCorreosProcesoLlenadoArchivoBaseWeb.ListarPorProceso(_db, "ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB", __anio, __mes, __version, __idAgencia).OrderByDescending(r => r.IdUnico).Where(v => v.IdStatusProceso == 1).ToList();

                        if (consultaInfoProceso.Any())
                        {
                            numReportes = consultaInfoProceso.FirstOrDefault().IdTipoProceso;
                            idUsuario = consultaInfoProceso.FirstOrDefault().IdUsuario;
                        }
                        else
                        {
                        
                            int idAgencia = Convert.ToInt32(idAgenciaYusuario[0]);
                            Agencia agencia = Agencia.Buscar(_db, idAgencia);

                            int idUsuarioGenerador = Convert.ToInt32(idAgenciaYusuario[1]);

                            nombreArchivo = $"{nombreBase}_{idUsuarioGenerador}_{idAgencia}__{DateTime.Now:yy-MM-dd}.txt";
                            msgErr = $"No se encontraron procesos activos con los siguientes parametros  {__siglasPestania} - {__anio}  {__mes}  {__version}  {__idAgencia}";
                            escribeLog.WriteLog(msgErr, 14066); 
                        }
                    }

                    List<string> pestanias = new List<string> { "BG", "OPL", "PX", "RMAN", "RM", "Cuentas", "VP", "AG", "CC", "AF", "SI", "SC",
                    "GANMT", "GANMF", "GANMVI", "GGNM", "GASM", "GSM", "GHPM", "GRM", "GDM" }; //"D"

                    if (args.Length == 7)
                    {
                        if (args[3] == "TODOS")
                        {
                            #region Agencias
                            string[] idAgenciaYusuario_ = args[5].Split('_');

                            var idUser = Convert.ToInt32(idAgenciaYusuario_[1]);
                            __claveUnica = Convert.ToInt32(args[6]);

                            int idAgencia = Convert.ToInt32(idAgenciaYusuario_[0]);
                            Agencia agencia = Agencia.Buscar(_db, idAgencia);

                            #region Agencias Adicionales
                            if (agencia.Id == 588)
                                agencia.Siglas = "AFAC";
                            else if (agencia.Id == 596)
                                agencia.Siglas = "DUC";
                            else if (agencia.Id == 286)
                                agencia.Siglas = "LYM";
                            else if (agencia.Id == 212)
                                agencia.Siglas = "EAH";
                            else if (agencia.Id == 348)
                                agencia.Siglas = "EAP";
                            else if (agencia.Id == 255)
                                agencia.Siglas = "EAZ";
                            else if (agencia.Id == 646)
                                agencia.Siglas = "LAR";
                            else if (agencia.Id == 585)
                                agencia.Siglas = "LMP";
                            else if (agencia.Id == 300)
                                agencia.Siglas = "MMB";
                            else if (agencia.Id == 365)
                                agencia.Siglas = "MMC";
                            else if (agencia.Id == 366)
                                agencia.Siglas = "MMR";
                            else if (agencia.Id == 373)
                                agencia.Siglas = "MMAB";

                            #endregion
                            #endregion

                            //Empieza Procesamiento
                            int numErrors = 0;
                            var consuLoteReportesAll = EnvioCorreosProcesoLlenadoArchivoBaseWeb.ListarPorClaveUnica(_db, __claveUnica, idUser);

                            if (consuLoteReportesAll.Any()) //Update status a 2 - En Proceso
                            {
                                var consuReporteStatusIniciado = consuLoteReportesAll.Where(t => t.IdStatusProceso == 1).ToList();
                                if (consuReporteStatusIniciado.Any())
                                {
                                    consuReporteStatusIniciado.First().IdStatusProceso = 2;
                                    consuReporteStatusIniciado.First().DATEUPDAT = DateTime.Now.Date;
                                    consuReporteStatusIniciado.First().TIMEUPDAT = DateTime.Now.TimeOfDay;
                        
                                    _db.Update(idUser, __numPrograma, consuReporteStatusIniciado.FirstOrDefault());
                                }

                                foreach (string pestania in pestanias)
                                {
                                    try
                                    {
                                        ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db, Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[4].ToString(), pestania, agencia.Nombre);
                                        proceso.LlenaExcel();
                                        //Actualiza status reporte individual
                                        EnvioCorreosProcesoLlenadoArchivoBaseWeb statProceso = EnvioCorreosProcesoLlenadoArchivoBaseWeb.BuscarEspecifico(_db, agencia.Id, idUsuario, __claveUnica, 1, pestanias.Count(), Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), pestania, __version);
                                        if (statProceso != null)
                                        {
                                            statProceso.IdStatusProceso = 3;
                                            statProceso.DATEUPDAT = DateTime.Now;
                                            statProceso.TIMEUPDAT = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                                            _db.Update(idUser, __numPrograma, statProceso);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        numErrors++;
                                        EnvioCorreosProcesoLlenadoArchivoBaseWeb statProceso = EnvioCorreosProcesoLlenadoArchivoBaseWeb.BuscarEspecifico(_db, agencia.Id, idUsuario, __claveUnica, 1, pestanias.Count(), Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), pestania, __version);
                                        if (statProceso != null)
                                        {
                                            statProceso.IdStatusProceso = 2;
                                            statProceso.DATEUPDAT = DateTime.Now;
                                            statProceso.TIMEUPDAT = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                                            _db.Update(idUser, __numPrograma, statProceso);
                                        }
                                        Console.WriteLine($"Error al procesar la pestaña {pestania}: {ex.Message}");
                                    }
                                }
                            }

                            var consuReporteStatusEnProceso = consuLoteReportesAll.Where(t => t.IdStatusProceso == 2).ToList();
                            if (consuReporteStatusEnProceso.Any())
                            {
                                consuReporteStatusEnProceso.First().IdStatusProceso = 3;
                                consuReporteStatusEnProceso.First().DATEUPDAT = DateTime.Now.Date;
                                consuReporteStatusEnProceso.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                                _db.Update(idUser, __numPrograma, consuReporteStatusEnProceso.FirstOrDefault());
                                __TerminaProcesoTodos = true;
                            }

                            __esTodos = true;
                        }
                        else
                        {

                            __claveUnica = Convert.ToInt32(args[6]);
                            string[] idAgenciaYusuario = args[5].Split('_');
                            int idUsuarioGenerador = Convert.ToInt32(idAgenciaYusuario[1]);
                            UsuarioLog = idUsuarioGenerador;

                            List<EnvioCorreosProcesoLlenadoArchivoBaseWeb> consuRegistrosReportes = new List<EnvioCorreosProcesoLlenadoArchivoBaseWeb>();

                            int idAgencia = Convert.ToInt32(idAgenciaYusuario[0]);
                            Agencia agencia = Agencia.Buscar(_db, idAgencia);

                            #region Log
                            string version = args[4].ToString();

                            string msgInicial = "";
                            msgInicial = $" :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: " + "\n";
                            msgInicial = $" : INICIA PROCESO ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB  ID_PROCESO_BASE_DATOS: {__claveUnica} - ID_USUARIO GENERADOR: {idUsuarioGenerador} :" + "\n\n";


                            nombreArchivo = $"{nombreBase}_{idUsuarioGenerador}_{idAgencia}_{version}__{DateTime.Now:yy-MM-dd}.txt";
                            // nombreArchivo = $"{nombreBase}{DateTime.Now:yy-MM-dd}__{idAgencia}_{version}.txt";
                            logFileName = nombreArchivo;
                            logFilePath = Path.Combine(logFolderPath, logFileName);
                            escribeLog.WriteLog(msgInicial, UsuarioLog);
                            #endregion

                            #region Agencias Adicionales
                            if (agencia.Id == 588)
                                agencia.Siglas = "AFAC";
                            else if (agencia.Id == 596)
                                agencia.Siglas = "DUC";
                            else if (agencia.Id == 286)
                                agencia.Siglas = "LYM";
                            else if (agencia.Id == 212)
                                agencia.Siglas = "EAH";
                            else if (agencia.Id == 348)
                                agencia.Siglas = "EAP";
                            else if (agencia.Id == 255)
                                agencia.Siglas = "EAZ";
                            else if (agencia.Id == 646)
                                agencia.Siglas = "LAR";
                            else if (agencia.Id == 585)
                                agencia.Siglas = "LMP";
                            else if (agencia.Id == 300)
                                agencia.Siglas = "MMB";
                            else if (agencia.Id == 365)
                                agencia.Siglas = "MMC";
                            else if (agencia.Id == 366)
                                agencia.Siglas = "MMR";
                            else if (agencia.Id == 373)
                                agencia.Siglas = "MMAB";

                            #endregion

                            __siglasAgencia = agencia.Siglas;
                            __agenciaNombre = agencia.Nombre;

                            ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                    Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[4].ToString(), args[3].ToString(), agencia.Nombre);

                            TimeSpan timeOutSmartIT = TimeSpan.FromSeconds(100);
                            logErr = $" - Asigna valores de Argumentos Correctamente en metodo ProcesoDeLlenadoExcelDesdeWeb." + "\n\n";
                            escribeLog.WriteLog(Program.logErr, UsuarioLog);
                            #region Consultas a tabla
                            consuRegistrosReportes = EnvioCorreosProcesoLlenadoArchivoBaseWeb.ListarPorClaveUnica(_db, __claveUnica, idUsuarioGenerador);
                            if (consuRegistrosReportes.Any())
                            {
                                consuRegistrosReportes = consuRegistrosReportes.Where(u =>
                                u.AnioParam == __anio &&
                                u.MesParam == __mes &&
                                u.VersionReporte == __version &&
                                u.IdAgencia == __idAgencia &&
                                u.IdStatusProceso == 1).OrderByDescending(o => o.IdUnico).ToList();

                                Stopwatch stopwatch = new Stopwatch();
                                stopwatch.Start();

                                if (consuRegistrosReportes.Any())
                                {
                                    logErr = $"   - Obtiene Registros de reportes en base a Parametros" + "\n\n";
                                    __existRegProceso = true;
                                    logErr = $"  - Inicia Proceso LlenaExcel con lo reportes:" + "\n";
                                    escribeLog.WriteLog(Program.logErr, UsuarioLog);
                                    try
                                    {
                                        #region Actualiza Status 2 en tabla  STATUS PROCESOS (FNDSTPRF)
                                        var validaStatusTablaStatusProcesos = StatusProcesos.BuscarPorClaveUnica(_db, __claveUnica.ToString());
                                        if (validaStatusTablaStatusProcesos != null)
                                        {
                                            try
                                            {
                                                if (validaStatusTablaStatusProcesos.IdStatusProceso == 1)
                                                {
                                                    _db.BeginTransaction();
                                                    validaStatusTablaStatusProcesos.IdStatusProceso = 2;
                                                    validaStatusTablaStatusProcesos.DATEUPDAT = DateTime.Now.Date;
                                                    validaStatusTablaStatusProcesos.TIMEUPDAT = DateTime.Now.TimeOfDay;
                                                    validaStatusTablaStatusProcesos.DescStatusProceso = "En Proceso...";

                                                    _db.Update(idUsuarioGenerador, 1499, validaStatusTablaStatusProcesos);

                                                    _db.CommitTransaction();
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                _db.RollbackTransaction();
                                            }
                                        }
                                        #endregion

                                        proceso.LlenaExcel(consuRegistrosReportes);

                                        logErr = $"  ---- Sin errores en metodo proceso.LlenaExcel(listReport); " + "\n\n";
                                        escribeLog.WriteLog(Program.logErr, UsuarioLog);
                                    }
                                    catch (Exception ex)
                                    {
                                        logErr = $"  - Error en proceso.LlenaExcel(list); {ex.Message} - {ex.StackTrace}"  + "\n\n";
                                        escribeLog.WriteLog(Program.logErr, UsuarioLog);                                     
                                    }                                 
                                }
                                else
                                {
                                    logErr = $"  - No hay reportes para este proceso con el STATUS 1 (Iniciado)" + "\n";
                                    escribeLog.WriteLog(Program.logErr, UsuarioLog);
                                }

                                stopwatch.Stop();
                                var tiempoTotal = stopwatch.Elapsed;
                                if (tiempoTotal >= timeOutSmartIT)
                                {
                                    //Envia Correo
                                }
                            }
                            else
                            {
                                Console.WriteLine("No hay registros en BD con los valores capturados: " + __claveUnica + " - " + idUsuarioGenerador);
                                logErr = $"   - No hay registros en BD con los valores capturados:  + { __claveUnica} - { idUsuarioGenerador}" + "\n\n";
                                escribeLog.WriteLog(Program.logErr, UsuarioLog);
                            }
                            #endregion
                        }

                    }
                    else if (args.Length == 6)
                    {
                        string[] idAgenciaYusuario = args[5].Split('_');

                        int idAgencia = Convert.ToInt32(idAgenciaYusuario[0]);
                        Agencia agencia = Agencia.Buscar(_db, idAgencia);

                        if (agencia.Id == 588)
                            agencia.Siglas = "AFAC";
                        else if (agencia.Id == 596)
                            agencia.Siglas = "DUC";
                        else if (agencia.Id == 286)
                            agencia.Siglas = "LYM";
                        else if (agencia.Id == 212)
                            agencia.Siglas = "EAH";
                        else if (agencia.Id == 348)
                            agencia.Siglas = "EAP";
                        else if (agencia.Id == 255)
                            agencia.Siglas = "EAZ";
                        else if (agencia.Id == 646)
                            agencia.Siglas = "LAR";
                        else if (agencia.Id == 585)
                            agencia.Siglas = "LMP";
                        else if (agencia.Id == 300)
                            agencia.Siglas = "MMB";
                        else if (agencia.Id == 365)
                            agencia.Siglas = "MMC";
                        else if (agencia.Id == 366)
                            agencia.Siglas = "MMR";
                        else if (agencia.Id == 373)
                            agencia.Siglas = "MMAB";

                        ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[4].ToString(), args[3].ToString(), agencia.Nombre);

                        TimeSpan timeOutSmartIT = TimeSpan.FromSeconds(100);

                        Stopwatch stopwatch = new Stopwatch();
                        stopwatch.Start();

                        proceso.LlenaExcel(consultaInfoProceso);


                        stopwatch.Stop();
                        var tiempoTotal = stopwatch.Elapsed;
                        if (tiempoTotal >= timeOutSmartIT)
                        {
                            //Envia Correo
                        }

                    }
                    else if (args.Length == 5)
                    {

                        string[] idAgenciaYusuario_ = args[4].Split('_');

                        var idUser = Convert.ToInt32(idAgenciaYusuario_[1]);

                        int idAgencia = Convert.ToInt32(idAgenciaYusuario_[0]);
                        Agencia agencia = Agencia.Buscar(_db, idAgencia);

                        #region Agencias Adicionales
                        if (agencia.Id == 588)
                            agencia.Siglas = "AFAC";
                        else if (agencia.Id == 596)
                            agencia.Siglas = "DUC";
                        else if (agencia.Id == 286)
                            agencia.Siglas = "LYM";
                        else if (agencia.Id == 212)
                            agencia.Siglas = "EAH";
                        else if (agencia.Id == 348)
                            agencia.Siglas = "EAP";
                        else if (agencia.Id == 255)
                            agencia.Siglas = "EAZ";
                        else if (agencia.Id == 646)
                            agencia.Siglas = "LAR";
                        else if (agencia.Id == 585)
                            agencia.Siglas = "LMP";
                        else if (agencia.Id == 300)
                            agencia.Siglas = "MMB";
                        else if (agencia.Id == 365)
                            agencia.Siglas = "MMC";
                        else if (agencia.Id == 366)
                            agencia.Siglas = "MMR";
                        else if (agencia.Id == 373)
                            agencia.Siglas = "MMAB";

                        #endregion

                        var consuStatusProcesosReportesAll = EnvioCorreosProcesoLlenadoArchivoBaseWeb.ListarPorProceso(_db, "ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB", Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), "", args[3], idAgencia, idUser).OrderByDescending(r => r.IdUnico).Where(v => v.IdStatusProceso == 1).ToList();

                        if (consuStatusProcesosReportesAll.Any())
                        {
                            consuStatusProcesosReportesAll.FirstOrDefault().IdStatusProceso = 2;
                            consuStatusProcesosReportesAll.FirstOrDefault().DATEUPDAT = DateTime.Now.Date;
                            consuStatusProcesosReportesAll.FirstOrDefault().TIMEUPDAT = DateTime.Now.TimeOfDay;
                            _db.Update(idUser, 9907, consuStatusProcesosReportesAll.FirstOrDefault());
                        }
                        foreach (string pestania in pestanias)
                        {
                            ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[3].ToString(), pestania, agencia.Nombre);
                            proceso.LlenaExcel();
                        }

                        #region Actualiza Proceso a Status 3 - Terminado
                        consuStatusProcesosReportesAll.FirstOrDefault().IdStatusProceso = 3;
                        consuStatusProcesosReportesAll.FirstOrDefault().DATEUPDAT = DateTime.Now.Date;
                        consuStatusProcesosReportesAll.FirstOrDefault().TIMEUPDAT = DateTime.Now.TimeOfDay;
                        _db.Update(idUser, 9907, consuStatusProcesosReportesAll.FirstOrDefault());
                        #endregion

                        #region EnviaCorreo
                        var correoUsuario_ = "";
                        var consuStatusProcesosReportesTerminados = EnvioCorreosProcesoLlenadoArchivoBaseWeb.ListarPorProceso(_db, "ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB", Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), "", args[3], idAgencia, idUser).OrderByDescending(r => r.IdUnico).Where(v => v.IdStatusProceso == 3).ToList();

                        if (consuStatusProcesosReportesTerminados.Any())
                        {
                            var buscaCorreoPorID_Usuario = Usuario.ListarPorIdunico(_db, idUsuario);

                            if (buscaCorreoPorID_Usuario != null && idUsuario != 0)
                            {
                                correoUsuario_ = buscaCorreoPorID_Usuario.MailPrimario != "" ? buscaCorreoPorID_Usuario.MailPrimario : buscaCorreoPorID_Usuario.MailSecundario;
                            }
                            else
                            {
                                Console.WriteLine("Id Usuario :" + idUsuario + " sin Correo ");
                            }

                            List<string> registros = new List<string>();
                            registros.Add(correoUsuario_);


                            EnviaCorreos env = new EnviaCorreos();

                            env.enviarCorreo(registros, "Proceso Llenado Info Web todos los reportes ha Terminado", "Reportes Procesados Terminados  :  ", null);
                        }
                        #endregion

                    }
                    else
                    {
                        List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1);
                        List<int> aIdAgencias = new List<int>();
                        aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                        List<Agencia> agencias = Agencia.ListarPorIds(_db, aIdAgencias);

                        foreach (Agencia agencia in agencias)
                        {
                            if (agencia.Id == 588)
                                agencia.Siglas = "AFAC";
                            else if (agencia.Id == 596)
                                agencia.Siglas = "DUC";
                            else if (agencia.Id == 286)
                                agencia.Siglas = "LYM";
                            else if (agencia.Id == 212)
                                agencia.Siglas = "EAH";
                            else if (agencia.Id == 348)
                                agencia.Siglas = "EAP";
                            else if (agencia.Id == 255)
                                agencia.Siglas = "EAZ";
                            else if (agencia.Id == 646)
                                agencia.Siglas = "LAR";
                            else if (agencia.Id == 585)
                                agencia.Siglas = "LMP";

                            foreach (string pestania in pestanias)
                            {
                                ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                    Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[3].ToString(), pestania, agencia.Nombre);
                                proceso.LlenaExcel();
                            }
                        }
                    }

                    if (args[3] == "TODOS")
                    {
                        if (__esTodos && __TerminaProcesoTodos)
                        {

                            //Validacion Correo de Usuario
                            string mailUsuario = "";
                            var buscaCorreoPorID_Usuario = Usuario.ListarPorIdunico(_db, idUsuario);
                            string nomUsuario = buscaCorreoPorID_Usuario.NombreReal;

                            if (buscaCorreoPorID_Usuario != null && idUsuario != 0)
                            {
                                mailUsuario = buscaCorreoPorID_Usuario.MailPrimario != "" ? buscaCorreoPorID_Usuario.MailPrimario : buscaCorreoPorID_Usuario.MailSecundario;
                            }
                            else
                            {
                                Console.WriteLine("Id Usuario :" + idUsuario + " sin Correo ");
                            }
                            var reportesGenerados = new string[] { };
                            foreach (var itemR in pestanias)
                            {
                                reportesGenerados = reportesGenerados.Concat(new string[] { itemR }).ToArray();
                            }
                            string registrosSeparadosPorComas = string.Join(", ", reportesGenerados);

                            //Manda correo
                            List<string> registrosReportes = new List<string>();
                            //mailUsuario = "emjuarez@grupoautofin.com";
                            registrosReportes.Add(mailUsuario);

                            List<string> fileNames = new List<string>();

                            EnviaCorreos env = new EnviaCorreos();

                            var resMail = Program.RutaExcelArchivoBaseCorreo; //Valor de variable Global - Ruta Excel Procesado
                            if (resMail != "")
                            {
                                fileNames.Add(resMail);
                                env.enviarCorreo(registrosReportes, "Proceso Llenado Info Web Todos los Reportes Terminado", "Reportes Procesados Terminados  :  " + registrosSeparadosPorComas, fileNames);
                            }
                            else
                            {
                                env.enviarCorreo(registrosReportes, "Proceso Llenado Info Web Todos los Reportes Terminado", "Reportes Procesados Terminados  :  " + registrosSeparadosPorComas, null);
                            }

                            Console.WriteLine("----------------- Se Envió Email al Usuario: " + nomUsuario + "  IdUsuario: " + buscaCorreoPorID_Usuario.Id + "\n\n");
                            StatusProcesos finalizaProceso = StatusProcesos.BuscarPorClaveUnica(_db, Convert.ToString(__claveUnica));
                            if (finalizaProceso != null)
                            {
                                finalizaProceso.IdStatusProceso = 3;
                                finalizaProceso.DescStatusProceso = "Terminado";
                                _db.Update(idUsuario, __numPrograma, finalizaProceso);
                                Console.WriteLine("----------------- Se Actualiza Status del proceso, clave unica " + __claveUnica + "\n\n");
                            }
                            Thread.Sleep(1000);
                        }
                        else
                        {
                            Console.WriteLine("Error en el Proceso de Generación de Todo los Reportes");
                        }
                    }
                    else //Reportes individuales
                    {
                        #region Valida Reportes con Status 3 Teminados
                        logErr = $"  Empieza Validacion de Reportes con Status 3 Terminado " + "\n\n";
                        escribeLog.WriteLog(Program.logErr, UsuarioLog);

                        int numReportesLote = 0;
                        string mailUsuario = "";
                        #region Obtiene num total de reportes
                        int countReportes = 0;
                        var ListaReportesArchivoBaseWeb = ParametrosProcesosFina.BuscarPorNombre(_db, "Reportes Archivo Base Web");
                        if (ListaReportesArchivoBaseWeb != null)
                        {
                            var stringIds = ListaReportesArchivoBaseWeb.ValorParametro;
                            List<int> intIdsList = stringIds.Split(',').Select(int.Parse).ToList();
                             countReportes = intIdsList.Count();
                        }
                        else
                        {
                            logErr = $"  --No existe registro de ids de reportes en tabla de parametros--  " + "\n\n";
                            escribeLog.WriteLog(Program.logErr, UsuarioLog);
                        }
                            #endregion
                        int numeroReportesTotales = countReportes; //Ajustar dependiendo el numero de reportes totales
                        bool esTodos = false;
                        if (__existRegProceso) //Proceso con 7 argumentos
                        {
                            var consuRegReportes = EnvioCorreosProcesoLlenadoArchivoBaseWeb.ListarPorClaveUnica(_db, __claveUnica).OrderByDescending(x => x.IdUnico).ToList();
                            if (consuRegReportes.Any())
                            {
                                if (consuRegReportes.Count == numeroReportesTotales)
                                {
                                    if (consuRegReportes.FirstOrDefault().IdTipoProceso == numeroReportesTotales)
                                    {
                                        esTodos = true;
                                    }
                                }

                            }

                            if (!consuRegReportes.Any())
                            {
                                Program.logErr = $"Sin registros en consulta consuRegReportes:";
                                escribeLog.WriteLog(Program.logErr, UsuarioLog);
                            }
                          

                            if (consuRegReportes.Any())
                            {
                                if (true) //  !esTodos Comparacion del total de reportes vs los registros encontrados en bd con la misma clave de lote de reportes
                                {
                                    //ajustar para cuando sea uno o mas reportes y cuando sea TODOS
                                    numReportesLote = consuRegReportes.Count();
                                    var consuRegReporCompletados = consuRegReportes.Where(x => x.IdStatusProceso == 3).ToList();
                                    var consuRegReporProcesando = consuRegReportes.Where(x => x.IdStatusProceso != 3).ToList();

                                    if (consuRegReporCompletados.Count == numReportesLote) //Lote de Reportes Completado - status 3
                                    {

                                        #region Ajuste Gastos
                                        try
                                        {
                                            Console.WriteLine("Ajusta pestañas gastos inicio!");
                                            Program.logErr = "Ajusta pestañas gastos inicio!";
                                            escribeLog.WriteLog(Program.logErr, UsuarioLog);

                                            pestanias = new List<string> { "GANMT", "GANMF", "GANMVI", "GGNM", "GASM", "GSM", "GHPM", "GRM", "GDM", "RM" }; //"D"

                                            foreach (string pestania in pestanias)
                                            {
                                                ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                                    Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), __idAgencia, __siglasAgencia, args[4].ToString(), pestania, __agenciaNombre);
                                                proceso.LlenaExcelAjustaDiferencia();

                                                Console.WriteLine("Pestaña de gastos " + pestania);
                                                Program.logErr = "Pestaña de gastos " + pestania;
                                                escribeLog.WriteLog(Program.logErr, UsuarioLog);
                                            }

                                            Console.WriteLine("Ajusta pestañas gastos fin!");
                                            Program.logErr = "Ajusta pestañas gastos fin!";
                                            escribeLog.WriteLog(Program.logErr, UsuarioLog);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine("Ajusta pestañas gastos, ex." + ex.Message);
                                            Program.logErr = "Ajusta pestañas gastos, ex." + ex.Message;
                                            escribeLog.WriteLog(Program.logErr, UsuarioLog);
                                        }
                                        #endregion

                                        #region Update Status 3 en Tabla Status Procesos - FNDSTPRF
                                        if (true)
                                        {
                                            try
                                            {
                                                var validaStatus3_TablaStatusProcesos = StatusProcesos.BuscarPorClaveUnica(_db, __claveUnica.ToString());
                                                if (validaStatus3_TablaStatusProcesos != null)
                                                {
                                                    if (validaStatus3_TablaStatusProcesos.IdStatusProceso != 3)
                                                    {
                                                        _db.BeginTransaction();

                                                        validaStatus3_TablaStatusProcesos.IdStatusProceso = 3;
                                                        validaStatus3_TablaStatusProcesos.DATEUPDAT = DateTime.Now.Date;
                                                        validaStatus3_TablaStatusProcesos.TIMEUPDAT = DateTime.Now.TimeOfDay;
                                                        validaStatus3_TablaStatusProcesos.DescStatusProceso = "Terminado";
                                                        _db.Update(idUsuario, 1499, validaStatus3_TablaStatusProcesos);

                                                        _db.CommitTransaction();
                                                    }
                                                }
                                                else
                                                {
                                                    Program.logErr = $"No se encuentra clave unica {__claveUnica} para actualizar a status 3 en StatusProcesos";
                                                    escribeLog.WriteLog(Program.logErr, UsuarioLog);
                                                }                                                                                                                           
                                            }
                                            catch (Exception ex)
                                            {
                                                Program.logErr = $"Erro al intentar hace Update a StatusProceso (status 3) {ex.Message} {ex.StackTrace}";
                                                escribeLog.WriteLog(Program.logErr, UsuarioLog);

                                                _db.RollbackTransaction();
                                            }
                                        }
                                        #endregion

                                        var reportesGenerados = new string[] { };
                                        foreach (var itemR in consuRegReporCompletados)
                                        {
                                            reportesGenerados = reportesGenerados.Concat(new string[] { itemR.SiglasReporte }).ToArray();
                                        }
                                        string registrosSeparadosPorComas = string.Join(", ", reportesGenerados);

                                        //Validacion Correo de Usuario
                                        var buscaCorreoPorID_Usuario = Usuario.ListarPorIdunico(_db, idUsuario);
                                        string nomUsuario = buscaCorreoPorID_Usuario.NombreReal;

                                        if (buscaCorreoPorID_Usuario != null && idUsuario != 0)
                                        {
                                            mailUsuario = buscaCorreoPorID_Usuario.MailPrimario != "" ? buscaCorreoPorID_Usuario.MailPrimario : buscaCorreoPorID_Usuario.MailSecundario;
                                        }
                                        else
                                        {
                                            Console.WriteLine("Id Usuario :" + idUsuario + " sin Correo ");
                                            Program.logErr = $"Id Usuario :" + idUsuario + " sin Correo ";
                                            escribeLog.WriteLog(Program.logErr, UsuarioLog);
                                        }

                                        escribeLog.WriteLog(Program.logErr, UsuarioLog);

                                        //Manda correo
                                        List<string> registrosReportes = new List<string>();
                                        registrosReportes.Add(mailUsuario);

                                        List<string> fileNames = new List<string>();

                                        EnviaCorreos env = new EnviaCorreos();

                                        var resMail = Program.RutaExcelArchivoBaseCorreo; //Valor de variable Global - Ruta Excel Procesado
                                        if (resMail != "")
                                        {
                                            fileNames.Add(resMail);
                                            env.enviarCorreo(registrosReportes, " -Proceso Llenado Info Web Terminado", "Reportes Procesados Terminados  :  " + registrosSeparadosPorComas, fileNames);
                                        }
                                        else
                                        {
                                            env.enviarCorreo(registrosReportes, "Proceso Llenado Info Web Terminado", "Reportes Procesados Terminados  :  " + registrosSeparadosPorComas, null);
                                        }

                                        Console.WriteLine("----------------- Se Envió Email al Usuario: " + nomUsuario + "  IdUsuario: " + buscaCorreoPorID_Usuario.Id + "\n\n\n");
                                        Program.logErr = $"-------- Se Envió Email al Usuario: " + nomUsuario + "  IdUsuario: " + buscaCorreoPorID_Usuario.Id + "\n\n\n";
                                        escribeLog.WriteLog(Program.logErr, UsuarioLog);
                                        Thread.Sleep(1000);

                                    }
                                    else
                                    {
                                        //Reportes del lote faltantes por terminar
                                    }
                                }


                            }

                        }
                        else
                        {
                            logErr = $"  __Erro en : __existRegProceso ... " + "\n\n";
                            escribeLog.WriteLog(Program.logErr, UsuarioLog);
                        }
                        #endregion
                    }

                    #endregion
                }
                else if (args[0] == "ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_AJUSTA_DIFERENCIA_REPORTES_GASTOS") //ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_AJUSTA_DIFERENCIA_REPORTES_GASTOS 2024 11 V1 O ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_AJUSTA_DIFERENCIA_REPORTES_GASTOS 2024 11 V1 28 O ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_AJUSTA_DIFERENCIA_REPORTES_GASTOS 2024 11 GGNM V1 28
                {
                    #region ARCHIVO_BASE_EXCEL_LLENADO_DESDE_WEB_AJUSTA_DIFERENCIA_REPORTES_GASTOS

                    List<string> pestanias = new List<string> { "GANMT", "GANMF", "GANMVI", "GGNM", "GASM", "GSM", "GHPM", "GRM", "GDM", "RM" }; //"D"

                    if (args.Length == 6)
                    {
                        int idAgencia = Convert.ToInt32(args[5]);
                        Agencia agencia = Agencia.Buscar(_db, idAgencia);

                        if (agencia.Id == 588)
                            agencia.Siglas = "AFAC";
                        else if (agencia.Id == 596)
                            agencia.Siglas = "DUC";
                        else if (agencia.Id == 286)
                            agencia.Siglas = "LYM";
                        else if (agencia.Id == 212)
                            agencia.Siglas = "EAH";
                        else if (agencia.Id == 348)
                            agencia.Siglas = "EAP";
                        else if (agencia.Id == 255)
                            agencia.Siglas = "EAZ";
                        else if (agencia.Id == 646)
                            agencia.Siglas = "LAR";
                        else if (agencia.Id == 585)
                            agencia.Siglas = "LMP";
                        else if (agencia.Id == 300)
                            agencia.Siglas = "MMB";
                        else if (agencia.Id == 365)
                            agencia.Siglas = "MMC";
                        else if (agencia.Id == 366)
                            agencia.Siglas = "MMR";
                        else if (agencia.Id == 373)
                            agencia.Siglas = "MMAB";

                        nombreArchivo = $"{nombreBase}_0_{agencia.Id}__{DateTime.Now:yy-MM-dd}.txt";

                        ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[4].ToString(), args[3].ToString(), agencia.Nombre);
                        proceso.LlenaExcelAjustaDiferencia();
                    }
                    else if (args.Length == 5)
                    {
                        int idAgencia = Convert.ToInt32(args[4]);
                        Agencia agencia = Agencia.Buscar(_db, idAgencia);

                        if (agencia.Id == 588)
                            agencia.Siglas = "AFAC";
                        else if (agencia.Id == 596)
                            agencia.Siglas = "DUC";
                        else if (agencia.Id == 286)
                            agencia.Siglas = "LYM";
                        else if (agencia.Id == 212)
                            agencia.Siglas = "EAH";
                        else if (agencia.Id == 348)
                            agencia.Siglas = "EAP";
                        else if (agencia.Id == 255)
                            agencia.Siglas = "EAZ";
                        else if (agencia.Id == 646)
                            agencia.Siglas = "LAR";
                        else if (agencia.Id == 585)
                            agencia.Siglas = "LMP";
                        else if (agencia.Id == 300)
                            agencia.Siglas = "MMB";
                        else if (agencia.Id == 365)
                            agencia.Siglas = "MMC";
                        else if (agencia.Id == 366)
                            agencia.Siglas = "MMR";
                        else if (agencia.Id == 373)
                            agencia.Siglas = "MMAB";

                        nombreArchivo = $"{nombreBase}_0_{agencia.Id}__{DateTime.Now:yy-MM-dd}.txt";

                        foreach (string pestania in pestanias)
                        {
                            ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[3].ToString(), pestania, agencia.Nombre);
                            proceso.LlenaExcelAjustaDiferencia();
                        }
                    }
                    else
                    {
                        List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1);
                        List<int> aIdAgencias = new List<int>();
                        aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                        List<Agencia> agencias = Agencia.ListarPorIds(_db, aIdAgencias);

                        foreach (Agencia agencia in agencias)
                        {
                            if (agencia.Id == 588)
                                agencia.Siglas = "AFAC";
                            else if (agencia.Id == 596)
                                agencia.Siglas = "DUC";
                            else if (agencia.Id == 286)
                                agencia.Siglas = "LYM";
                            else if (agencia.Id == 212)
                                agencia.Siglas = "EAH";
                            else if (agencia.Id == 348)
                                agencia.Siglas = "EAP";
                            else if (agencia.Id == 255)
                                agencia.Siglas = "EAZ";
                            else if (agencia.Id == 646)
                                agencia.Siglas = "LAR";
                            else if (agencia.Id == 585)
                                agencia.Siglas = "LMP";

                            nombreArchivo = $"{nombreBase}_0_{agencia.Id}__{DateTime.Now:yy-MM-dd}.txt";

                            foreach (string pestania in pestanias)
                            {
                                ProcesoDeLlenadoExcelDesdeWeb proceso = new ProcesoDeLlenadoExcelDesdeWeb(_db,
                                    Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), agencia.Id, agencia.Siglas, args[3].ToString(), pestania, agencia.Nombre);
                                proceso.LlenaExcel();
                            }
                        }
                    }

                    #endregion
                }
                else if (args[0] == "ARCHIVO_BASE_EXCEL_GASTO_CORPORATIVO") //ARCHIVO_BASE_EXCEL_GASTO_CORPORATIVO 2024 12 Diciembre
                {
                    #region ARCHIVO_BASE_EXCEL_GASTO_CORPORATIVO

                    nombreArchivo = $"{nombreBase}_0_ARCHIVO_BASE_EXCEL_GASTO_CORPORATIVO__{DateTime.Now:yy-MM-dd}.txt";

                    ProcesoDeLlenadoGastoCorporativo proceso = new ProcesoDeLlenadoGastoCorporativo(_db,
                                Convert.ToInt32(args[1]), Convert.ToInt32(args[2]), args[3].ToString());
                    proceso.LecturaExcel();

                    #endregion
                }
                else if (args[0] == "ARCHIVO_BASE_LLENA_SI") //ARCHIVO_BASE_LLENA_SI 2024 11 O ARCHIVO_BASE_LLENA_SI 2024 11 26
                {
                    # region ARCHIVO_BASE_LLENA_SI

                    if (args.Length == 4)
                    {
                        int idAgencia = Convert.ToInt32(args[4]);
                        Agencia agencia = Agencia.Buscar(_db, idAgencia);

                        if (agencia.Id == 588)
                            agencia.Siglas = "AFAC";
                        else if (agencia.Id == 596)
                            agencia.Siglas = "DUC";
                        else if (agencia.Id == 286)
                            agencia.Siglas = "LYM";
                        else if (agencia.Id == 212)
                            agencia.Siglas = "EAH";
                        else if (agencia.Id == 348)
                            agencia.Siglas = "EAP";
                        else if (agencia.Id == 255)
                            agencia.Siglas = "EAZ";
                        else if (agencia.Id == 646)
                            agencia.Siglas = "LAR";
                        else if (agencia.Id == 585)
                            agencia.Siglas = "LMP";
                        else if (agencia.Id == 300)
                            agencia.Siglas = "MMB";
                        else if (agencia.Id == 365)
                            agencia.Siglas = "MMC";
                        else if (agencia.Id == 366)
                            agencia.Siglas = "MMR";
                        else if (agencia.Id == 373)
                            agencia.Siglas = "MMAB";

                        ProcesoDeLlenadoSI proceso = new ProcesoDeLlenadoSI(_db,
                                Convert.ToInt32(args[4]), Convert.ToInt32(args[2]), Convert.ToInt32(args[3]));
                    
                    }

                    #endregion
                }
                else //MES-1 BLOQUE1 28     O      //MES-1 BLOQUE1
                {
                    #region MES-1 BLOQUE1 28
                                        
                    /* 
                        Es necesario cambiar los BATS a producción para que la información sea obtenida de las tablas de Producción y sea la información correcta
                        Validar si la fecha de hoy = fecha Cierre Archivo Base
                    */

                    PeriodoContable fechaArchivoBase = new PeriodoContable();

                    DateTime fecha_Actual = DateTime.Today.AddMonths(-1);

                    if (args[0] == "ACTUAL")
                        fecha_Actual = DateTime.Now;
                    else if (args[0].Contains("MES-"))
                    {
                        char[] sep = { '-' };
                        string[] mesesMenos = args[0].Split(sep);
                        fecha_Actual = DateTime.Today.AddMonths(Convert.ToInt32(mesesMenos[1]) * -1);
                    }

                    int idAgencia = 0;

                    if (args.Length == 3) //MES-1 BLOQUE1 28
                    {
                        if (args[2] != "")
                            idAgencia = Convert.ToInt32(args[2]);
                    }

                    PeriodoContable Fecha = PeriodoContable.FechaLimiteArchivoBase(_db, fecha_Actual.Year, fecha_Actual.Month);
                    // A la fecha de entrega del archivo Base agregarle un día más, para que tome ese día como último!
                    // A partir del siguiente día comenzará el proceso del mes anterior!

                    //if (DateTime.Now.Date == Fecha.FechaCierre.AddDays(3).Date)
                    //{
                    //System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Sistemas DVA\errorlogProcesoTabArchivoBase.txt");
                    //file.AutoFlush = true;

                    DateTime hoy = fecha_Actual;
                    //int anio = (hoy.Month == 1) ? hoy.Year - 1 : hoy.Year;
                    int anio = hoy.Year;
                    //Iniciar el proceso de llenado
                    for (int i = anio; i <= anio; i++)
                    {
                        Console.WriteLine("Inicia Año: " + i);
                        List<AgenciasReportes> LiAgenciasReportes = new List<AgenciasReportes>();

                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == 342);

                        /*LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == 32 ||
                        a.IdAgencia == 126 ||
                        a.IdAgencia == 113 ||
                        a.IdAgencia == 116);*/

                        //pendiente la agencia 202
                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia <= 126);

                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia != 202);

                        LiAgenciasReportes = AgenciasReportes.Listar(_db, 1).Ordenar("IdAgencia", false).DistinctList();
                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == 275);
                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == 32 || a.IdAgencia == 113 || a.IdAgencia == 116 || a.IdAgencia == 126);
                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == 98 || a.IdAgencia == 131 || a.IdAgencia == 275 || a.IdAgencia == 115 || a.IdAgencia == 138);
                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == 26);
                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == 88 || a.IdAgencia == 95);
                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == 42);
                        //LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == 42 || a.IdAgencia == 205 || a.IdAgencia == 38 || a.IdAgencia == 191 
                        //    || a.IdAgencia == 39 || a.IdAgencia == 581 || a.IdAgencia == 581);

                        if (args.Length == 3)
                            LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia == idAgencia);
                        else
                        {
                            if (args[1] == "BLOQUE1")
                            {
                                LiAgenciasReportes = AgenciasReportes.Listar(_db, 1).Ordenar("IdAgencia", false).DistinctList();
                                LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia <= 38);
                            }
                            else if (args[1] == "BLOQUE2")
                            {
                                LiAgenciasReportes = AgenciasReportes.Listar(_db, 1).Ordenar("IdAgencia", false).DistinctList();
                                LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia > 38 && a.IdAgencia <= 190);
                            }
                            else if (args[1] == "BLOQUE3")
                            {
                                LiAgenciasReportes = AgenciasReportes.Listar(_db, 1).Ordenar("IdAgencia", false).DistinctList();
                                LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia > 190 && a.IdAgencia <= 275);
                            }
                            else if (args[1] == "BLOQUE4")
                            {
                                LiAgenciasReportes = AgenciasReportes.Listar(_db, 1).Ordenar("IdAgencia", false).DistinctList();
                                LiAgenciasReportes = LiAgenciasReportes.FindAll(a => a.IdAgencia > 275 && a.IdAgencia <= 594);
                            }
                        }

                        Console.WriteLine("Total Agencias: " + LiAgenciasReportes.Count);

                        #region Generar informacion V1 y V2 Archivo Base

                        //int mes = (hoy.Month == 1) ? 12 : hoy.Month - 1;
                        int mesInicio = hoy.Month;
                        int mesFin = DateTime.Now.Month;

                        if (mesFin == 1)
                            mesFin = (hoy.Month == 1) ? 12 : hoy.Month;

                        if (mesInicio == 12)
                            mesFin = 12;

                        for (int x = mesInicio; x <= mesFin; x++)
                        //for (int x = 1; x <= mes; x++)
                        {
                            foreach (AgenciasReportes aAgenciasReportes in LiAgenciasReportes)
                            {
                                nombreArchivo = $"{nombreBase}_0_{aAgenciasReportes.IdAgencia}__{DateTime.Now:yy-MM-dd}.txt";

                                _db.DbCnx.CerrarConexion();
                                _db.DbCnx.AbrirConexion();

                                Console.WriteLine("Mes: " + x + " Del Año " + i);

                                var elimina = EliminaConcepto(_db, aAgenciasReportes.IdAgencia, x, i);
                                var eliminaExtralibros = EliminaConceptoExtralibros(_db, aAgenciasReportes.IdAgencia, x, i);
                                try
                                {
                                    // Eliminar de la tabla el Mes. año. Agencia a depositar para evitar duplicidades

                                    try
                                    {
                                        //LLENAR TABLA RESULTADOS MENSUALES   FNDRESMEN V1 Y V2
                                        _db.BeginTransaction();
                                        ProcesoDeLLenadoTablaRM proceso = new ProcesoDeLLenadoTablaRM(_db, aAgenciasReportes.IdAgencia, x, i);
                                        proceso.InsertaRegistros();
                                        _db.CommitTransaction();
                                    }
                                    catch (Exception e)
                                    {
                                        //file.WriteLine("***************Error RM*******************************************");
                                        //file.WriteLine(e.Message);
                                        continue;
                                    }

                                    try
                                    {
                                        /* Proceso llenado tabla VENTAS para reportes Aspectos Generales */
                                        var eliminaAg = EliminaConceptotablaMST(_db, aAgenciasReportes.IdAgencia, x, i);
                                        _db.BeginTransaction();
                                        ProcesoDeLLenadoTablaArchivoBase ag = new ProcesoDeLLenadoTablaArchivoBase(_db, aAgenciasReportes.IdAgencia, "Todos", x, i);
                                        ag.InsertaRegistrosActualesParaAgencia();
                                        _db.CommitTransaction();
                                    }
                                    catch (Exception e)
                                    {
                                        //file.WriteLine("***************Error AG*******************************************");
                                        //file.WriteLine(e.Message);
                                        continue;
                                    }

                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("[ERROR][Main-All][Agencia=" + aAgenciasReportes.Id + "]" + " Error: " + e);
                                    _db.RollbackTransaction();
                                }
                            }

                            _db.DbCnx.CerrarConexion();
                            _db.DbCnx.AbrirConexion();

                            //var eliminaTmp = EliminaConceptosTemporales(_db, aAgenciasReportes.IdAgencia);
                        }

                        #endregion
                    }
                    //file.Dispose();
                    //file.Close();
                    //}

                    #endregion
                }



                //escribeLog.WriteLog(Program.logErr,UsuarioLog);

                #region Guarda Log
                if (true)
                {
                    Program.logErr = $"**______ TERMINA PROCESO ______**" + "\n\n\n\n";
                    escribeLog.WriteLog(Program.logErr, 14066);
                }
               
                #endregion

            }
            catch (Exception ex)
            {
                Program escribeLog = new Program();
                Program.logErr = $" - Error en Program: : {ex.Message} - {ex.StackTrace}" + "\n\n";
                escribeLog.WriteLog(Program.logErr, 14066);
 
            }
         
        }

        public void WriteLog(string Valor, int aIdUsuario)
        {
            try
            {

                string nombreArchivo_ = logFolderPath + nombreArchivo;
                FileInfo fi = new FileInfo(nombreArchivo_);

                if (!Directory.Exists(fi.DirectoryName))
                    Directory.CreateDirectory(fi.DirectoryName);

                StreamWriter sw = new StreamWriter(nombreArchivo_, true);
                sw.WriteLine(DateTime.Now + "| " + Valor);
                sw.Flush();
                sw.Close();
                sw.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine("----------------- Error al escribir Log:  " + ex.Message  + "\n\n\n");
            }

        }

        public static int EliminaConcepto(DB2Database aDB, int aIdAgencia, int aMes, int aAnio)
        {
            var query = @"delete FROM [PREFIX]FINA.FNDRESMEN
                    WHERE FIFNIDCIAU = " + aIdAgencia + @"
                    AND FIFNYEAR = " + aAnio + @"
                    AND FIFNMONTH = " + aMes + @"
                    AND FIFNCPTD NOT IN (1031, 1032, 1033) ";
            Console.WriteLine("Eliminando registros anteriores RM");
            var registrosEliminados = aDB.SetQuery(query);
            return registrosEliminados;
        }

        public static int EliminaConceptoExtralibros(DB2Database aDB, int aIdAgencia, int aMes, int aAnio)
        {
            var query = @"delete FROM [PREFIX]FINA.FNDRSMENE
                    WHERE FIFNIDCIAU = " + aIdAgencia + @"
                    AND FIFNYEAR = " + aAnio + @"
                    AND FIFNMONTH = " + aMes + @"
                    AND FIFNCPTD NOT IN (1031, 1032, 1033) ";
            Console.WriteLine("Eliminando registros anteriores RM Extralibros");
            var registrosEliminados = aDB.SetQuery(query);
            return registrosEliminados;
        }


        public static int EliminaConceptosTemporales(DB2Database aDB, int aIdAgencia)
        {
            var query = "CALL [PREFIX]FINA.SPDRESMENTMP (" + aIdAgencia + ")";
            Console.WriteLine("Eliminando registros Temporales utilizados para el cálculo");
            var registrosEliminados = aDB.SetQuery(query);
            return registrosEliminados;
        }


        public static int EliminaConceptotablaMST(DB2Database aDB, int aIdAgencia, int aMes, int aAnio)
        {
            var query = "CALL [PREFIX]FINA.SPDRPTMST (" + aIdAgencia + ", " + aMes + "," + aAnio + ")";
            Console.WriteLine("Eliminando registros anteriores");
            var registrosEliminados = aDB.SetQuery(query);
            return registrosEliminados;
        }


        public static int EliminaConceptotablaSC(DB2Database aDB, int aIdAgencia, int aMes, int aAnio)
        {
            var query = "CALL [PREFIX]FINA.SPDSITCAR (" + aIdAgencia + ", " + aMes + "," + aAnio + ")";
            Console.WriteLine("Eliminando registros anteriores");
            var registrosEliminados = aDB.SetQuery(query);
            return registrosEliminados;
        }

    }

    public class Ejecutor
    {
        DVAConstants.Constants constantes = new DVAConstants.Constants();
        //public DVAControls.csLog m_log = null;
        csUtilerias m_util = new csUtilerias();
        //Boolean escribeLOG = true;

        //public static List<csDetalleClase> dtDetalleClaseAtributosMetodos = new List<csDetalleClase>();
        //public static BindingList<csGECCLASETemplate> dtClases = new BindingList<csGECCLASETemplate>();
        //public static BindingList<csGECCUEMATemplate> dtCuentasMayor = new BindingList<csGECCUEMATemplate>();
        //public static BindingList<csCOCATCTSTemplate> dtCuentasContables = new BindingList<csCOCATCTSTemplate>();
        //public static BindingList<csCODCFGPRTemplate> dtConfigProceso = new BindingList<csCODCFGPRTemplate>();
        //public static BindingList<csGEDATRPRTemplate> dtAtributosProceso = new BindingList<csGEDATRPRTemplate>();
        //public static BindingList<csGECPRXMOTemplate> dtProcesos = new BindingList<csGECPRXMOTemplate>();

        //public static DataTable dtGECCLASE = new DataTable();
        //public static DataTable dtCOCATCTS = new DataTable();
        //public static DataTable dtGEDATRPR = new DataTable();

        const Int32 PARAM_RUTAXML = 0;

        private Int32 idPoliza;

        public static String strRutaXML = "";

        public Int32 IdPoliza
        {
            get { return idPoliza; }
            set { idPoliza = value; }
        }

        public void EjecutaProceso(string[] args)
        {
            if (Program.m_log == null)
            {
                if (Program.escribeLOG)
                {
                    //if (Program.m_log == null)
                    Program.m_log = new DVAControls.csLog("Demonio_PGC_EjecutaProceso", constantes.CA_GE_DIRECTORIO_LOGS, Convert.ToInt32(args[1]));
                }
            }

            DateTime inicio = DateTime.Now;

            try
            {
                //String[] configuraciones = File.ReadAllLines(ProcesoGeneracionContabilidad.Properties.Settings.Default.RUTA_CONFIG + "\\" 
                //String[] configuraciones = File.ReadAllLines(getRutaSmartIT() + "\\"
                //    + ProcesoDeLLenadoTablaArchivoBase.Properties.Settings.Default.ARCHIVO_CONFIGURACION);
                //strRutaXML = configuraciones[PARAM_RUTAXML].Substring(configuraciones[PARAM_RUTAXML].IndexOf("=") + 1);

                //if (!Directory.Exists(strRutaXML))
                //    return;

                //Program m_start = new Program();
                //m_start.escribeLOG = Program.escribeLOG;
                //m_start.m_log = Program.m_log;                
                //m_start.Inicia(args);
                Program.Main(args);
            }
            catch (Exception ex)
            {
                if (Program.m_log != null)
                {
                    if (Program.escribeLOG)
                    {
                        Program.m_log.AgregaRegistro("ex.Message: " + ex.Message);
                        Program.m_log.AgregaRegistro("ex.InnerException: " + ex.InnerException);
                        Program.m_log.AgregaRegistro("ex.StackTrace: " + ex.StackTrace);
                    }
                }
            }

            DateTime fin = DateTime.Now;
            TimeSpan dif = fin - inicio;
            Console.WriteLine("[SEGUNDOS][" + dif.TotalSeconds + "]");

            if (Program.m_log != null)
            {
                if (Program.escribeLOG)
                    Program.m_log.AgregaRegistro("[SEGUNDOS][" + dif.TotalSeconds + "]");
            }
        }

        public static String getRutaSmartIT()
        {
            String ruta = Application.ExecutablePath;
            ruta = ruta.Substring(0, ruta.LastIndexOf("\\"));

            if (ruta.Contains("c:\\windows"))
                ruta = "C:\\AutosystV2";

            return ruta;
        }
    }
}
