using DVADB;
using DVAModelsReflection;
using DVAModelsReflection.Models.AUSA;
using DVAModelsReflection.Models.CONT;
using DVAModelsReflection.Models.GRAL;
using DVAModelsReflection.Models.NOM;
using DVAModelsReflection.Models.PSI;
using DVAModelsReflection.Models.REFA.Interfaces;
using DVAModelsReflection.Models.TESO;
using DVAModelsReflectionFINA.Models.FINA;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static System.Collections.Specialized.BitVector32;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    public class ProcesoDeLlenadoExcelDesdeWeb
    {
        DB2Database _db = null;
        Missing Sin = Missing.Value;
        ProcesoGeneraComparacionExcelVsWeb proc;
        Program escribeLog = new Program();
        string query = "";

        [DllImport("user32")]
        public static extern int GetWindowThreadProcessId(int hwnd, ref int lpdwProcessId);

        //private string Program.logFolderPath;
        //private string logFileName;
        //private string Program.logFilePath;
        // string Program.logFilePath = @"\\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase\LogProcesoConsola";
       
        


        String[] alfabeto = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
        public String[] ALFABETO
        {
            get { return this.alfabeto; }
        }

        int idAgencia = 0;
        int anio = 0;
        int mes = 0;
        string siglasAgencia = "";
        string siglas = "";
        string razonSocial = "";
        string version = "";
        string pestania = "";

        //string ruta = @"C:\Users\fevangelista\Desktop\ARCHIVOS BASE EXCEL\";
        //string archivoOrigenV1 = @"C:\Users\fevangelista\Desktop\ARCHIVOS BASE EXCEL\";
        //string archivoOrigenV2 = @"C:\Users\fevangelista\Desktop\ARCHIVOS BASE EXCEL\";

        //string ruta = Program.RutaServidor;
        //string archivoOrigenV1 = Program.RutaServidor;
        //string archivoOrigenV2 = Program.RutaServidor;

        //string ruta = @"D:\Users\jasoria\Desktop\Escritorio\PMO\Proyectos 2023\ARCHIVO BASE\ARCHIVOS BASE EXCEL\";
        //string archivoOrigenV1 = @"D:\Users\jasoria\Desktop\Escritorio\PMO\Proyectos 2023\ARCHIVO BASE\ARCHIVOS BASE EXCEL\";
        //string archivoOrigenV2 = @"D:\Users\jasoria\Desktop\Escritorio\PMO\Proyectos 2023\ARCHIVO BASE\ARCHIVOS BASE EXCEL\";

        //string ruta = @"C:\Users\dgaytan\Desktop\";
        //string archivoOrigenV1 = @"C:\Users\dgaytan\Desktop\";
        //string archivoOrigenV2 = @"C:\Users\dgaytan\Desktop\";

        //string ruta = @"C:\Users\malopez\Desktop\";
        //string archivoOrigenV1 = @"C:\Users\malopez\Desktop\";
        //string archivoOrigenV2 = @"C:\Users\malopez\Desktop\";

        string ruta = @"\\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase\";
        string archivoOrigenV1 = @"\\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase\";
        string archivoOrigenV2 = @"\\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase\";

        //string ruta = @"\\10.5.2.118\fina\ARCHIVO_BASE\";
        //string archivoOrigenV1 = @"\\10.5.2.118\fina\ARCHIVO_BASE\";
        //string archivoOrigenV2 = @"\\10.5.2.118\fina\ARCHIVO_BASE\";

        // "\\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase\2024\5\V1\ARCHIVO_BASE_WEB"
        // \\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase

        //string ruta = @"E:\AutosystV2\ProcesoLlenadoDeTablaArchivoBase\";
        //string archivoOrigenV1 = @"E:\AutosystV2\ProcesoLlenadoDeTablaArchivoBase\";
        //string archivoOrigenV2 = @"E:\AutosystV2\ProcesoLlenadoDeTablaArchivoBase\";

        //string ruta = @"E:\ARCHIVOS BASE EXCEL\";


        DataRow[] drBGWebV1 = null;
        DataRow[] drBGWebV1Acum = null;
        DataRow[] drBGWebV2 = null;
        DataRow[] drBGWebV2Acum = null;

        public ProcesoDeLlenadoExcelDesdeWeb(DB2Database _db, int aAnio, int aMes, int aIdAgencia, string aSiglasAgencia, string aVersion, string aPestania, string aRazonSocial)
        {
            this._db = _db;
            this.idAgencia = aIdAgencia;
            this.siglasAgencia = aSiglasAgencia;
            this.mes = aMes;
            this.anio = aAnio;
            this.version = aVersion;
            //this.siglas = aSiglas;
            this.razonSocial = aRazonSocial;
            this.pestania = aPestania;
            ruta = ruta + anio + "\\" + mes + "\\" + version + "\\" + "ARCHIVO_BASE_WEB" + "\\";
            archivoOrigenV1 = archivoOrigenV1 + anio + "\\" + mes + "\\" + version + "\\" + "ARCHIVO_BASE_EXCEL" + "\\";
            archivoOrigenV2 = archivoOrigenV2 + anio + "\\" + mes + "\\" + version + "\\" + "ARCHIVO_BASE_EXCEL" + "\\";
            Console.WriteLine("[INICIA LECTURA ARCHIVO BASE EXCEL PESTAÑA " + pestania + "]: ");
            Console.WriteLine("[ID_AGENCIA]: " + idAgencia);
            Console.WriteLine("[AÑO]: " + anio);
            Console.WriteLine("[MES]: " + mes);
            Console.WriteLine("[VERSION]: " + version);
            Console.WriteLine("[SIGLAS]: " + siglas);
            Console.WriteLine("[PESTAÑA]: " + pestania);

            proc = new ProcesoGeneraComparacionExcelVsWeb(_db, aAnio, aMes, aSiglasAgencia);
        }

        public void LlenaExcel(List<EnvioCorreosProcesoLlenadoArchivoBaseWeb> listReportes)
        {
            int errorExcel = 0;
            int _mExcelPID = 0;
            int IdProceso = 0;
            int IdProcesoLog = 0;
            bool excelAbierto = false;

            //String Program.logErr = "";
            // Program.logFolderPath = @"\\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase\LogProcesoConsola";
            //string nombreBase = "error_log_";

            //string Program.nombreArchivo = $"{nombreBase}{DateTime.Now:yy-MM-dd}__{idAgencia}_{version}.txt";

            //logFileName = Program.nombreArchivo;
            //Program.logFilePath = Path.Combine(Program.logFolderPath, logFileName);
            Program.logErr = $" -\\ Empieza Proceso de Metodo LLena Excel del reporte: --_ {pestania} _--" + "\n\n\n";
            escribeLog.WriteLog(Program.logErr, 14066);
            try
            {
                if (version == "V1")
                {
                    ruta = ruta + "ARCHIVO BASE " + anio + " " + siglasAgencia + ".xls";
                    archivoOrigenV1 = archivoOrigenV1 + "ARCHIVO BASE " + anio + " " + siglasAgencia + ".xls";
                    if (!File.Exists(ruta))
                        File.Copy(archivoOrigenV1, ruta);

                    Program.RutaExcelArchivoBaseCorreo = ruta;

                    try
                    {
                        Program.logErr = $" Valida si el archivo esta siendo utilizado por otro proceso o esta bloqueado por otro proceso!" + "\n\n";
                        escribeLog.WriteLog(Program.logErr, 14066);
                        using (FileStream fs = new FileStream(ruta, FileMode.Open, FileAccess.Read))
                        { Program.logErr = $" - El Archivo: {ruta}  no esta siendo utilizado por otro proceso se continua con el proceso" + "\n\n"; escribeLog.WriteLog(Program.logErr, 14066); }
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine("El archivo de Excel está siendo utilizado por otro proceso");

                        Console.WriteLine("Error: " + ex.Message);
                        Program.logErr = "--- Error_ El archivo de Excel está siendo utilizado por otro proceso:   " + ex.Message + "\n";
                        escribeLog.WriteLog(Program.logErr, 14066);
                        excelAbierto = true;

                        var res = CloseExcelWorkbook(ruta);
                        Program.logErr = res;
                        escribeLog.WriteLog(Program.logErr, 14066);
                    }
                }
                else
                {
                    ruta = ruta + "ARCHIVO BASE " + anio + " VERSION 2 " + siglasAgencia + ".xls";
                    archivoOrigenV2 = archivoOrigenV2 + "ARCHIVO BASE " + anio + " VERSION 2 " + siglasAgencia + ".xls";
                    if (!File.Exists(ruta))
                        File.Copy(archivoOrigenV2, ruta);

                    Program.RutaExcelArchivoBaseCorreo = ruta;


                    try
                    {
                        using (FileStream fs = new FileStream(ruta, FileMode.Open, FileAccess.Read))
                        { Program.logErr = $" - El Archivo: {ruta}  no esta siendo utilizado por otro proceso se continua con el proceso" + "\n\n"; escribeLog.WriteLog(Program.logErr, 14066); }
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine("El archivo de Excel está siendo utilizado por otro proceso");
                        Console.WriteLine("Error: " + ex.Message);
                        Program.logErr = "--- Error_ El archivo de Excel está siendo utilizado por otro proceso:   " + ex.Message + "\n";
                        escribeLog.WriteLog(Program.logErr, 14066);
                        excelAbierto = true;

                       var res =  CloseExcelWorkbook(ruta);
                        Program.logErr = res;
                        escribeLog.WriteLog(Program.logErr, 14066);
                    }
                }

                if (excelAbierto)
                {
                    throw new InvalidOperationException("Excel abierto en otro proceso.");
                }

                Program.logErr = $" - Se empieza a abrir la instancia de Excel: " + "\n\n";
                escribeLog.WriteLog(Program.logErr, 14066);
                ExcelApp.Application app = new ExcelApp.Application();
                //ExcelApp.Workbook Libro = app.Workbooks.Open(ruta, Sin, false, Sin, Sin, Sin, Sin, Sin, Sin, Sin, false, Sin, Sin, Sin, Sin);
                ExcelApp.Workbook Libro = app.Workbooks.Open(ruta, 0, false, 5, "", "", false, ExcelApp.XlPlatform.xlWindows, "", true, true, 0, true, false, false);

                var numReportes = Libro.Worksheets.Count;

                var idUser = 0;
                var numPrograma = 101010;
                bool esEnvioCorreo = true;
                int countReportesProcesados = 0;
                string msgError = "";
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                Program.logErr = $"------ Empieza Iteración de los reportes para su lectura y escritura ----- " + "\n";
                escribeLog.WriteLog(Program.logErr, 14066);
                foreach (ExcelApp.Worksheet ws in Libro.Worksheets)
                {
                    if ((ws.Name == pestania) && (pestania == "D"))
                        LlenaPestaniaD(ws);
                    else if ((ws.Name == pestania) && (pestania == "RM"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;


                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma , filtraPorSiglasRepo.First());
                            Program.logErr = $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                            escribeLog.WriteLog(Program.logErr, 14066);
                        }
                        try
                        {
                            LlenaPestaniaRM(ws);
                        }
                        catch (Exception ex) {
                            msgError = ex.Message;

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr = $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        escribeLog.WriteLog(Program.logErr, 14066);
                        countReportesProcesados++;

                    }
                    else if ((ws.Name == pestania) && (pestania == "RMAN"))
                    {
                        try
                        {
                            var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();


                            if (filtraPorSiglasRepo.Any())
                            {
                                _db.BeginTransaction();
                                idUser = filtraPorSiglasRepo.First().IdUsuario;
                                filtraPorSiglasRepo.First().IdStatusProceso = 2;
                                filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                                filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                                _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                                _db.CommitTransaction();
                                Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                            }

                            LlenaPestaniaRMANyRMAU(ws);

                            filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                            filtraPorSiglasRepo.First().IdStatusProceso = 3;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.BeginTransaction();
                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            _db.CommitTransaction();
                            Program.logErr += $" Se actualiza el STATUS a 3(Terminado) del Reporte {pestania}" + "\n";
                            countReportesProcesados++;
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Console.WriteLine("Error en reporte: " + " _ " + msgError);
                            _db.RollbackTransaction();
                        }
                    }
                    else if ((ws.Name == pestania) && (pestania == "OPL"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaOPL(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                       
                    else if ((ws.Name == pestania) && (pestania == "PX"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }
                        else
                        {
                            Console.WriteLine("El reporte aun no esta en status Iniciado");
                        }


                        try
                        {
                            LlenaPestaniaPX(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                       
                    else if ((ws.Name == pestania) && (pestania == "BG"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaBGV2(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                        
                    else if ((ws.Name == pestania) && (pestania == "Cuentas"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaCuentas(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                        
                    else if ((ws.Name == pestania) && (pestania == "VP"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaVP(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en  Reporte {pestania}  Error: {ex.Message}" + "\n";

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                        
                    else if ((ws.Name == pestania) && (pestania == "AG"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaAG(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en Reporte {pestania} :  {ex.Message}" + "\n";

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                    else if ((ws.Name == pestania) && (pestania == "CC"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaAG(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en Reporte {pestania} :  {ex.Message}" + "\n";

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }

                    else if ((ws.Name == pestania) && (pestania == "AF"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaAF(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en Reporte {pestania} -  {ex.Message} " + "\n";
                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                       
                    else if ((ws.Name == pestania) && (pestania == "SI"))
                    {
                        try
                        {
                            var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                            
                            if (filtraPorSiglasRepo.Any())
                            {
                                _db.BeginTransaction();
                                idUser = filtraPorSiglasRepo.First().IdUsuario;
                                filtraPorSiglasRepo.First().IdStatusProceso = 2;
                                filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                                filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                                _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                                _db.CommitTransaction();
                                Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                            }

                            LlenaPestaniaSI(ws);

                            filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                            filtraPorSiglasRepo.First().IdStatusProceso = 3;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.BeginTransaction();
                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            _db.CommitTransaction();
                            Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                            countReportesProcesados++;
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Console.WriteLine("Error en reporte: " + " _ " + msgError);
                            Program.logErr += $" Error en Reporte {pestania}   {ex.Message}" + "\n";
                            _db.RollbackTransaction();
                        }                    
                    }
                       
                    else if ((ws.Name == pestania) && (pestania == "SC"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaSC(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en Reporte {pestania} -  {ex.Message}" + "\n";

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                        
                    else if ((ws.Name == pestania) && (pestania == "GANMT"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaGANMT(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" ERROr en Reporte {pestania} -  {ex.Message}" + "\n";
                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }

                    else if ((ws.Name == pestania) && (pestania == "GANMF"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaGANMF(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" ERROR en Reporte {pestania} - {ex.Message}" + "\n";

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                        
                    else if ((ws.Name == pestania) && (pestania == "GANMVI"))
                    {
                        bool eError = false;
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaGANMVI(ws);
                        }
                        catch (Exception ex)
                        {
                            eError = true;
                            msgError = ex.Message;
                            Program.logErr += $" Error en Reporte {pestania} - {ex.Message}" + "\n";
                        }

                        if (!eError)
                        {
                            filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                            filtraPorSiglasRepo.First().IdStatusProceso = 3;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                            countReportesProcesados++;
                        }
                       
                    }
                       

                    else if ((ws.Name == pestania) && (pestania == "GGNM"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaGGNM(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en Reporte {pestania} - {ex.Message}" + "\n";

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                       
                    else if ((ws.Name == pestania) && (pestania == "GASM"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaGASM(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en  Reporte {pestania} - {ex.Message}" + "\n";
                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                       
                    else if ((ws.Name == pestania) && (pestania == "GSM"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaGSM(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" ErroR Reporte {pestania}   - {ex.Message}" + "\n";

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                       
                    else if ((ws.Name == pestania) && (pestania == "GHPM"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaGHPM(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en Reporte {pestania}   - {ex.Message}" + "\n";
                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                        
                    else if ((ws.Name == pestania) && (pestania == "GRM"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaGRM(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en Reporte {pestania} - {ex.Message}" + "\n";

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                       
                    else if ((ws.Name == pestania) && (pestania == "GDM"))
                    {
                        var filtraPorSiglasRepo = listReportes.Where(r => r.SiglasReporte == pestania && r.IdStatusProceso == 1).ToList();

                        if (filtraPorSiglasRepo.Any())
                        {
                            idUser = filtraPorSiglasRepo.First().IdUsuario;
                            filtraPorSiglasRepo.First().IdStatusProceso = 2;
                            filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                            filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                            _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                            Program.logErr += $" Se actualiza el STATUS a 2(En Proceso) del Reporte {pestania}" + "\n";
                        }

                        try
                        {
                            LlenaPestaniaGDM(ws);
                        }
                        catch (Exception ex)
                        {
                            msgError = ex.Message;
                            Program.logErr += $" Error en Reporte {pestania}  - {ex.Message}" + "\n";

                        }

                        filtraPorSiglasRepo.Where(t => t.IdStatusProceso == 2).ToList();
                        filtraPorSiglasRepo.First().IdStatusProceso = 3;
                        filtraPorSiglasRepo.First().DATEUPDAT = DateTime.Now.Date;
                        filtraPorSiglasRepo.First().TIMEUPDAT = DateTime.Now.TimeOfDay;

                        _db.Update(idUser, numPrograma, filtraPorSiglasRepo.First());
                        Program.logErr += $" Se actualiza el STATUS a 3 del Reporte {pestania}" + "\n";
                        countReportesProcesados++;
                    }
                    else
                    {
                        Program.logErr += $" Error, Siglas Reporte {pestania}  - no Encontrado" + "\n";
                    }

                }
               
                stopwatch.Stop();
                var tiempoTranscurrido = stopwatch.Elapsed;

                if (tiempoTranscurrido <= TimeSpan.FromSeconds(150))
                {
                    //No envia Correo
                    esEnvioCorreo = false;
                }

                if (!esEnvioCorreo)
                {
                    EnviaCorreos env = new EnviaCorreos();
                }


                Libro.Save();
                Program.logErr = $" -Se guardan los cambios realizados en el Excel " + "\n";
                escribeLog.WriteLog(Program.logErr, 14066);
                Libro.Close();
                Program.logErr = $" -Se cierra correctamente  el Excel " + "\n";
                escribeLog.WriteLog(Program.logErr, 14066);

                Program.logErr = $" -Se procede a ralizar el app.Quit(); del Excel: " + "\n";
                escribeLog.WriteLog(Program.logErr, 14066);
                app.Quit();
                Program.logErr = $" -app.Quit(); Success " + "\n";
                escribeLog.WriteLog(Program.logErr, 14066);

            }
            catch (Exception ex)
            {
                if (ex.Message == "Excel abierto en otro proceso.")
                {
                    Program.logErr = "Excel abierto en otro proceso. Cerrar desde el Servidor! " + "\n\n";
                    escribeLog.WriteLog(Program.logErr, 14066);
                }
                else
                {
                    Console.WriteLine(ex.Message);
                    Program.logErr = "-- Error en el Reporte: " + pestania + "   " + ex.Message + "\n";
                    escribeLog.WriteLog(Program.logErr, 14066);
                    Program.logErr = "-- StackTrace: " + pestania + "   " + ex.StackTrace + "\n";
                    escribeLog.WriteLog(Program.logErr, 14066);

                    _mExcelPID = 0;
                    IdProceso = IdProcesoLog;
                    Program.logErr = $"-- Obtiene Valores instancia Excel: {IdProceso} " + "\n";
                    GetWindowThreadProcessId(IdProceso, ref _mExcelPID);
                    System.Diagnostics.Process proceso = Process.GetProcessById(_mExcelPID);
                    Program.logErr = "-- Obtiene Id de la instancia Excel para forzar cierre: " +  "\n";

                    proceso.Kill();
                    Program.logErr = "_--- Se forzó el cierre de la instancia de Excel y se mato con Exito ---_ " + "\n\n";
                    escribeLog.WriteLog(Program.logErr, 14066);
                }
            }

            #region Guarda Log
            //try
            //{
            //    if (Program.logErr != "")
            //    {
            //        if (File.Exists(Program.logFilePath))
            //        {

            //            using (StreamWriter writer = File.AppendText(Program.logFilePath))
            //            {
            //                writer.WriteLine($"[{DateTime.Now:HH:mm:ss}]  {Program.logErr}");
            //            }
            //        }
            //        else
            //        {
            //            string newLogFileName = $"log_{DateTime.Now:yyyyMMddHHmmss}.txt";
            //            string newLogFilePath = Path.Combine(Program.logFolderPath, Program.nombreArchivo);

            //            using (StreamWriter writer = File.CreateText(newLogFilePath))
            //            {
            //                writer.WriteLine($"[{DateTime.Now:HH:mm:ss}]  {Program.logErr}");
            //            }
            //        }
            //    }
             
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}

            #endregion
        }

        public void LlenaExcel()
        {
            int _mExcelPID = 0;
            int IdProceso = 0;
            int IdProcesoLog = 0;
            bool excelAbierto = false;

            //String Program.logErr = "";
            //Program.logFolderPath = @"\\SVP-TP2023\ProcesoLlenadoDeTablaArchivoBase\LogProcesoConsola";
            //string nombreBase = "error_log_";
            //string Program.nombreArchivo = $"{nombreBase}{DateTime.Now:yy-MM-dd}__{idAgencia}_{version}.txt";

            //logFileName = Program.nombreArchivo;
            //Program.logFilePath = Path.Combine(Program.logFolderPath, logFileName);

            Program.logErr += "-- Inicia Metodo LlenaExcel -- " + "\n";
            try
            {
                if (version == "V1")
                {
                    ruta = ruta + "ARCHIVO BASE " + anio + " " + siglasAgencia + ".xls";
                    archivoOrigenV1 = archivoOrigenV1 + "ARCHIVO BASE " + anio + " " + siglasAgencia + ".xls";
                    if (!File.Exists(ruta))
                        File.Copy(archivoOrigenV1, ruta);

                    try
                    {
                        using (FileStream fs = new FileStream(ruta, FileMode.Open, FileAccess.Read))
                        {
                            //Archivo no abierto
                        }
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine("El archivo de Excel está siendo utilizado por otro proceso");
                        Console.WriteLine("Error: " + ex.Message);
                        Program.logErr += "--- Error_ El archivo de Excel está siendo utilizado por otro proceso:   " + ex.Message + "\n";
                        excelAbierto = true;
                    }


                    Program.RutaExcelArchivoBaseCorreo = ruta;
                }
                else
                {
                    ruta = ruta + "ARCHIVO BASE " + anio + " VERSION 2 " + siglasAgencia + ".xls";
                    archivoOrigenV2 = archivoOrigenV2 + "ARCHIVO BASE " + anio + " VERSION 2 " + siglasAgencia + ".xls";
                    if (!File.Exists(ruta))
                        File.Copy(archivoOrigenV2, ruta);

                    try
                    {
                        using (FileStream fs = new FileStream(ruta, FileMode.Open, FileAccess.Read))
                        {
                           //Archivo no abierto en otro proceso
                        }
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine("El archivo de Excel está siendo utilizado por otro proceso");
                        Console.WriteLine("Error: " + ex.Message);
                        Program.logErr = "--- Error_ El archivo de Excel está siendo utilizado por otro proceso:   " + ex.Message + "\n";
                        excelAbierto = true;
                    }

                    string nombreProceso = Path.GetFileNameWithoutExtension(ruta);

                    Program.RutaExcelArchivoBaseCorreo = ruta;
                }


                if (excelAbierto)
                {
                    throw new InvalidOperationException("Excel abierto en otro proceso.");
                }
                ExcelApp.Application app = new ExcelApp.Application();
                IdProcesoLog = app.Hwnd;
                //ExcelApp.Workbook Libro = app.Workbooks.Open(ruta, Sin, false, Sin, Sin, Sin, Sin, Sin, Sin, Sin, false, Sin, Sin, Sin, Sin);
                ExcelApp.Workbook Libro = app.Workbooks.Open(ruta, 0, false, 5, "", "", false, ExcelApp.XlPlatform.xlWindows, "", true, true, 0, true, false, false);

                foreach (ExcelApp.Worksheet ws in Libro.Worksheets)
                {
                    if ((ws.Name == pestania) && (pestania == "D"))
                        LlenaPestaniaD(ws);
                    else if ((ws.Name == pestania) && (pestania == "RM"))
                        LlenaPestaniaRM(ws);
                    else if ((ws.Name == pestania) && (pestania == "RMAN"))
                        LlenaPestaniaRMANyRMAU(ws);
                    else if ((ws.Name == pestania) && (pestania == "OPL"))
                        LlenaPestaniaOPL(ws);
                    else if ((ws.Name == pestania) && (pestania == "PX"))
                        LlenaPestaniaPX(ws);
                    else if ((ws.Name == pestania) && (pestania == "BG"))
                        LlenaPestaniaBGV2(ws);
                    else if ((ws.Name == pestania) && (pestania == "Cuentas"))
                        LlenaPestaniaCuentas(ws);
                    else if ((ws.Name == pestania) && (pestania == "VP"))
                    {
                        try
                        {
                            LlenaPestaniaVP(ws);
                        }
                        catch (Exception e)
                        {

                        }
                    }
                    else if ((ws.Name == pestania) && (pestania == "AG"))
                        LlenaPestaniaAG(ws);
                    else if ((ws.Name == pestania) && (pestania == "CC"))
                        LlenaPestaniaCC(ws);
                    else if ((ws.Name == pestania) && (pestania == "AF"))
                        LlenaPestaniaAF(ws);
                    else if ((ws.Name == pestania) && (pestania == "SI"))
                        LlenaPestaniaSI(ws);
                    else if ((ws.Name == pestania) && (pestania == "SC"))
                        LlenaPestaniaSC(ws);
                    else if ((ws.Name == pestania) && (pestania == "GANMT"))
                        LlenaPestaniaGANMT(ws);
                    else if ((ws.Name == pestania) && (pestania == "GANMF"))
                        LlenaPestaniaGANMF(ws);
                    else if ((ws.Name == pestania) && (pestania == "GANMVI"))
                        LlenaPestaniaGANMVI(ws);
                    else if ((ws.Name == pestania) && (pestania == "GGNM"))
                        LlenaPestaniaGGNM(ws);
                    else if ((ws.Name == pestania) && (pestania == "GASM"))
                        LlenaPestaniaGASM(ws);
                    else if ((ws.Name == pestania) && (pestania == "GSM"))
                        LlenaPestaniaGSM(ws);
                    else if ((ws.Name == pestania) && (pestania == "GHPM"))
                        LlenaPestaniaGHPM(ws);
                    else if ((ws.Name == pestania) && (pestania == "GRM"))
                        LlenaPestaniaGRM(ws);
                    else if ((ws.Name == pestania) && (pestania == "GDM"))
                        LlenaPestaniaGDM(ws);
                }

                Libro.Save();
                Libro.Close();
               //bool seGuardo =  Libro.Saved;

                app.Quit();
                 _mExcelPID = 0;
                 IdProceso = app.Hwnd;
                GetWindowThreadProcessId(IdProceso, ref _mExcelPID);
                System.Diagnostics.Process proceso = Process.GetProcessById(_mExcelPID);
                proceso.Kill();
                Program.logErr = "---------  Termina Correctamente  ---------";
            }
            catch (Exception ex)
            {
                if (ex.Message == "Excel abierto en otro proceso.")
                {
                    Program.logErr = "Excel abierto en otro proceso. Cerrar desde el Servidor! " + "\n\n";
                }
                else
                {
                    Console.WriteLine(ex.Message);
                    Program.logErr = "-- Error en el Reporte: " + pestania + "   " + ex.Message + "\n";

                    _mExcelPID = 0;
                    IdProceso = IdProcesoLog;
                    GetWindowThreadProcessId(IdProceso, ref _mExcelPID);
                    System.Diagnostics.Process proceso = Process.GetProcessById(_mExcelPID);

                    proceso.Kill();
                    Program.logErr = "_--- Se forzó el cierre de la instancia de Excel ---_ " + "\n\n";
                }
               
            }

            #region Guarda Log
        
            #endregion
        }

        #region MyRegion
        static string CloseExcelWorkbook(string filePath)
        {
            Program escribeLog = new Program();
            string msg = "";
            ExcelApp.Application excelApp = null;
            try
            {
                excelApp = (ExcelApp.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException)
            {
                Console.WriteLine("No se encontró ninguna instancia de Excel en ejecución.");
                msg = "No se encontró ninguna instancia de Excel en ejecución." + "\n";
                escribeLog.WriteLog(msg, 14066);
                return msg;
            }

            bool workbookClosed = false;

            if (excelApp != null)
            {
                foreach (ExcelApp.Workbook workbook in excelApp.Workbooks)
                {
                    if (workbook.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook.Close(false); // Cerrar sin guardar cambios
                        workbookClosed = true;
                        Console.WriteLine($"El archivo {filePath} ha sido cerrado.");
                        msg = $"El archivo {filePath} ha sido cerrado." + "\n";
                        escribeLog.WriteLog(msg, 14066);
                        break;
                    }
                }

                if (!workbookClosed)
                {
                    Console.WriteLine($"El archivo {filePath} no se encontró en ninguna instancia de Excel.");
                    msg = $"El archivo {filePath} no se encontró en ninguna instancia de Excel." + "\n";
                    escribeLog.WriteLog(msg, 14066);
                    return msg;
                }

                int processId;
                GetWindowThreadProcessId(new IntPtr(excelApp.Hwnd), out processId);
                Process process = Process.GetProcessById(processId);
                try
                {
                    process.Kill();
                    Console.WriteLine($"Proceso de Excel con ID {processId} ha sido terminado.");
                    msg = $"Proceso de Excel con ID {processId} ha sido terminado." + "\n";
                    escribeLog.WriteLog(msg, 14066);
                    return msg;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al terminar el proceso de Excel con ID {processId}: {ex.Message}");
                    msg = $"Error al terminar el proceso de Excel con ID {processId}: {ex.Message}" + "\n";
                    escribeLog.WriteLog(msg, 14066);
                    return msg;
                }
            }

            return msg;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        #endregion

        public void LlenaExcelAjustaDiferencia()
        {
            try
            {
                if (version == "V1")
                    ruta = ruta + "ARCHIVO BASE " + anio + " " + siglasAgencia + ".xls";
                else
                    ruta = ruta + "ARCHIVO BASE " + anio + " VERSION 2 " + siglasAgencia + ".xls";

                ExcelApp.Application app = new ExcelApp.Application();
                //ExcelApp.Workbook Libro = app.Workbooks.Open(ruta, Sin, false, Sin, Sin, Sin, Sin, Sin, Sin, Sin, false, Sin, Sin, Sin, Sin);
                ExcelApp.Workbook Libro = app.Workbooks.Open(ruta, 0, false, 5, "", "", false, ExcelApp.XlPlatform.xlWindows, "", true, true, 0, true, false, false);

                foreach (ExcelApp.Worksheet ws in Libro.Worksheets)
                {
                    if ((ws.Name == pestania) && (pestania == "GANMT"))
                        AjustaPestaniasGastos(ws);
                    else if ((ws.Name == pestania) && (pestania == "GANMF"))
                        AjustaPestaniasGastos(ws);
                    else if ((ws.Name == pestania) && (pestania == "GANMVI"))
                        AjustaPestaniasGastos(ws);
                    else if ((ws.Name == pestania) && (pestania == "GGNM"))
                        AjustaPestaniasGastos(ws);
                    else if ((ws.Name == pestania) && (pestania == "GASM"))
                        AjustaPestaniasGastos(ws);
                    else if ((ws.Name == pestania) && (pestania == "GSM"))
                        AjustaPestaniasGastos(ws);
                    else if ((ws.Name == pestania) && (pestania == "GHPM"))
                        AjustaPestaniasGastos(ws);
                    else if ((ws.Name == pestania) && (pestania == "GRM"))
                        AjustaPestaniasGastos(ws);
                    else if ((ws.Name == pestania) && (pestania == "GDM"))
                        AjustaPestaniasGastos(ws);
                    else if ((ws.Name == pestania) && (pestania == "RM"))
                    {
                        if (version == "V2")
                            AjustaPestaniaRM(ws);
                    }
                }

                Libro.Save();
                Libro.Close();

                app.Quit();
                int _mExcelPID = 0;
                int IdProceso = app.Hwnd;
                GetWindowThreadProcessId(IdProceso, ref _mExcelPID);
                System.Diagnostics.Process proceso = Process.GetProcessById(_mExcelPID);
                proceso.Kill();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void LlenaPestaniaD(ExcelApp.Worksheet ws)
        {
            string celda = "";

            celda = "B4"; //INTRODUZCA LA CLAVE DE LA EMPRESA:  SIGLAS
            ws.get_Range(celda, celda).Formula = siglasAgencia;

            celda = "B6"; //ACTUALICE EL NUMERO DEL MES DEL AÑO ACTUAL:  MES
            ws.get_Range(celda, celda).Formula = mes;

            DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
            string nombreMes = formatoFecha.GetMonthName(mes).ToUpper();

            celda = "B9"; //ACTUALICE EL MES ACTUAL:  	MES DE MARZO 2024
            ws.get_Range(celda, celda).Formula = "MES DE " + nombreMes + " " + anio;

            celda = "B11"; //ACTUALICE EL MES ACTUAL:  	ACUMULADO A MARZO 2024
            ws.get_Range(celda, celda).Formula = "ACUMULADO A " + nombreMes + " " + anio;

            celda = "B13"; //ACTUALICE EL MES ACTUAL:  	ANALISIS DE RESULTADOS OPERATIVOS DEL MES DE MARZO 2024
            ws.get_Range(celda, celda).Formula = "ANALISIS DE RESULTADOS OPERATIVOS DEL MES DE " + nombreMes + " " + anio;

            celda = "B13"; //ACTUALICE EL MES ACTUAL:  	RESULTADOS DEL MES DE MARZO 2024
            ws.get_Range(celda, celda).Formula = "RESULTADOS DEL MES DE " + nombreMes + " " + anio;
        }

        public void LlenaPestaniaBG(ExcelApp.Worksheet ws)
        {
            List<ConceptosContables> conceptosV1 = ConceptosContables.ListarBGV1yV2(_db); //obtiene el listado de todos los conceptos contables
            //List<ConceptosContables> conceptosV2 = ConceptosContables.ListarBGV1yV2(_db);

            System.Data.DataTable dtBGWebV1 = (System.Data.DataTable)proc.GetBGWebV1();
            //System.Data.DataTable dtBGWebV2 = (System.Data.DataTable)proc.GetBGWebV2();

            //int MaxColumna = ws.Cells.SpecialCells(ExcelApp.XlCellType.xlCellTypeLastCell, ExcelApp.XlSpecialCellsValue.xlTextValues).Column;
            //int MaxRenglon = ws.Cells.SpecialCells(ExcelApp.XlCellType.xlCellTypeLastCell, ExcelApp.XlSpecialCellsValue.xlTextValues).Row;

            int r = 1;
            int vWeb = 0;
            string celda = "";

            //celda = "A1";

            //if (celda == "A1") //Razón Social
            //    ws.get_Range(celda, celda).Formula = razonSocial;

            //celda = "";

            foreach (ConceptosContables concepto in conceptosV1)
            {
                Reporte rep = new Reporte();

                r++;

                rep.ID_CONCEPTO = concepto.Id;
                rep.CONCEPTO = concepto.NombreConcepto;

                int cambioSigo = 1;
                if ((concepto.Id == 66) || (concepto.Id == 67) || (concepto.Id == 68) || (concepto.Id == 69) || (concepto.Id == 70) ||
                    (concepto.Id == 71) || (concepto.Id == 72) || (concepto.Id == 73) || (concepto.Id == 74) || (concepto.Id == 75) ||
                    (concepto.Id == 76) || (concepto.Id == 77) || (concepto.Id == 78) || (concepto.Id == 79) || (concepto.Id == 80) ||
                    (concepto.Id == 81) || (concepto.Id == 82) || (concepto.Id == 83) || (concepto.Id == 84) || (concepto.Id == 85))
                    cambioSigo = -1;

                if (version == "V1" || version == "V2")
                {
                    #region V1

                    if (idAgencia == 27)
                    {
                        drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                        drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 286)
                    {
                        drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                        drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 36)
                    {
                        drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                        drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 12)
                    {
                        drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                        drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 35)
                    {
                        drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                        drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 588)
                    {
                        drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                        drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 32)
                    {
                        drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                        drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 212)
                    {
                        drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                        drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else
                    {
                        drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                        drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                    }

                    #endregion
                }
                //else
                //{
                //    #region V2

                //    if (idAgencia == 27)
                //    {
                //        drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                //        drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                //    }
                //    else if (idAgencia == 286)
                //    {
                //        drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                //        drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                //    }
                //    else if (idAgencia == 36)
                //    {
                //        drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                //        drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                //    }
                //    else if (idAgencia == 12)
                //    {
                //        drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                //        drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                //    }
                //    else if (idAgencia == 35)
                //    {
                //        drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",595) AND ID_CONCEPTO = " + concepto.Id);
                //        drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",595) AND ID_CONCEPTO = " + concepto.Id);
                //    }
                //    else if (idAgencia == 588)
                //    {
                //        drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                //        drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                //    }
                //    else if (idAgencia == 32)
                //    {
                //        drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                //        drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                //    }
                //    else if (idAgencia == 212)
                //    {
                //        drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                //        drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                //    }
                //    else
                //    {
                //        drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                //        drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                //    }

                //    #endregion
                //}

                if (version == "V1" || version == "V2")
                {
                    #region V1

                    if (drBGWebV1.Length != 0)
                    {
                        vWeb = 0;

                        foreach (DataRow dr in drBGWebV1)
                        {
                            if (mes == 1)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]);
                            else if (mes == 2)
                                vWeb += Convert.ToInt32(dr["SALDO_FEBRERO"]);
                            else if (mes == 3)
                                vWeb += Convert.ToInt32(dr["SALDO_MARZO"]);
                            else if (mes == 4)
                                vWeb += Convert.ToInt32(dr["SALDO_ABRIL"]);
                            else if (mes == 5)
                                vWeb += Convert.ToInt32(dr["SALDO_MAYO"]);
                            else if (mes == 6)
                                vWeb += Convert.ToInt32(dr["SALDO_JUNIO"]);
                            else if (mes == 7)
                                vWeb += Convert.ToInt32(dr["SALDO_JULIO"]);
                            else if (mes == 8)
                                vWeb += Convert.ToInt32(dr["SALDO_AGOSTO"]);
                            else if (mes == 9)
                                vWeb += Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]);
                            else if (mes == 10)
                                vWeb += Convert.ToInt32(dr["SALDO_OCTUBRE"]);
                            else if (mes == 11)
                                vWeb += Convert.ToInt32(dr["SALDO_NOVIEMBRE"]);
                            else if (mes == 12)
                                vWeb += Convert.ToInt32(dr["SALDO_DICIEMBRE"]);

                            rep.WEB_V1 = vWeb * cambioSigo;
                        }

                        vWeb = 0;

                        foreach (DataRow dr in drBGWebV1Acum)
                        {
                            if (mes == 1)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]);
                            else if (mes == 2)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]);
                            else if (mes == 3)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]);
                            else if (mes == 4)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]);
                            else if (mes == 5)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                                    Convert.ToInt32(dr["SALDO_MAYO"]);
                            else if (mes == 6)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]);
                            else if (mes == 7)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                                    Convert.ToInt32(dr["SALDO_JULIO"]);
                            else if (mes == 8)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]);
                            else if (mes == 9)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]) +
                                    Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]);
                            else if (mes == 10)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]) +
                                    Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]) + Convert.ToInt32(dr["SALDO_OCTUBRE"]);
                            else if (mes == 11)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]) +
                                    Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]) + Convert.ToInt32(dr["SALDO_OCTUBRE"]) +
                                    Convert.ToInt32(dr["SALDO_NOVIEMBRE"]);
                            else if (mes == 12)
                                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]) +
                                    Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]) + Convert.ToInt32(dr["SALDO_OCTUBRE"]) +
                                    Convert.ToInt32(dr["SALDO_NOVIEMBRE"]) + Convert.ToInt32(dr["SALDO_DICIEMBRE"]);

                            rep.WEB_V1_ACUM = vWeb * cambioSigo;
                        }
                    }

                    #endregion
                }

                //rep.DIFF_V1 = rep.EXCEL_V1 - rep.WEB_V1;
                //rep.DIFF_V1_ACUM = rep.EXCEL_V1_ACUM - rep.WEB_V1_ACUM;

                //if (version == "V2")
                //{
                //    #region V2

                //    if (drBGWebV2.Length != 0)
                //    {
                //        vWeb = 0;

                //        foreach (DataRow dr in drBGWebV2)
                //        {
                //            if (mes == 1)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]);
                //            else if (mes == 2)
                //                vWeb += Convert.ToInt32(dr["SALDO_FEBRERO"]);
                //            else if (mes == 3)
                //                vWeb += Convert.ToInt32(dr["SALDO_MARZO"]);
                //            else if (mes == 4)
                //                vWeb += Convert.ToInt32(dr["SALDO_ABRIL"]);
                //            else if (mes == 5)
                //                vWeb += Convert.ToInt32(dr["SALDO_MAYO"]);
                //            else if (mes == 6)
                //                vWeb += Convert.ToInt32(dr["SALDO_JUNIO"]);
                //            else if (mes == 7)
                //                vWeb += Convert.ToInt32(dr["SALDO_JULIO"]);
                //            else if (mes == 8)
                //                vWeb += Convert.ToInt32(dr["SALDO_AGOSTO"]);
                //            else if (mes == 9)
                //                vWeb += Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]);
                //            else if (mes == 10)
                //                vWeb += Convert.ToInt32(dr["SALDO_OCTUBRE"]);
                //            else if (mes == 11)
                //                vWeb += Convert.ToInt32(dr["SALDO_NOVIEMBRE"]);
                //            else if (mes == 12)
                //                vWeb += Convert.ToInt32(dr["SALDO_DICIEMBRE"]);

                //            rep.WEB_V2 = vWeb * cambioSigo;
                //        }

                //        vWeb = 0;

                //        foreach (DataRow dr in drBGWebV2Acum)
                //        {
                //            if (mes == 1)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]);
                //            else if (mes == 2)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]);
                //            else if (mes == 3)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]);
                //            else if (mes == 4)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]);
                //            else if (mes == 5)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                //                    Convert.ToInt32(dr["SALDO_MAYO"]);
                //            else if (mes == 6)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                //                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]);
                //            else if (mes == 7)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                //                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                //                    Convert.ToInt32(dr["SALDO_JULIO"]);
                //            else if (mes == 8)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                //                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                //                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]);
                //            else if (mes == 9)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                //                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                //                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]) +
                //                    Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]);
                //            else if (mes == 10)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                //                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                //                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]) +
                //                    Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]) + Convert.ToInt32(dr["SALDO_OCTUBRE"]);
                //            else if (mes == 11)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                //                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                //                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]) +
                //                    Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]) + Convert.ToInt32(dr["SALDO_OCTUBRE"]) +
                //                    Convert.ToInt32(dr["SALDO_NOVIEMBRE"]);
                //            else if (mes == 12)
                //                vWeb += Convert.ToInt32(dr["SALDO_ENERO"]) + Convert.ToInt32(dr["SALDO_FEBRERO"]) +
                //                    Convert.ToInt32(dr["SALDO_MARZO"]) + Convert.ToInt32(dr["SALDO_ABRIL"]) +
                //                    Convert.ToInt32(dr["SALDO_MAYO"]) + Convert.ToInt32(dr["SALDO_JUNIO"]) +
                //                    Convert.ToInt32(dr["SALDO_JULIO"]) + Convert.ToInt32(dr["SALDO_AGOSTO"]) +
                //                    Convert.ToInt32(dr["SALDO_SEPTIEMBRE"]) + Convert.ToInt32(dr["SALDO_OCTUBRE"]) +
                //                    Convert.ToInt32(dr["SALDO_NOVIEMBRE"]) + Convert.ToInt32(dr["SALDO_DICIEMBRE"]);

                //            rep.WEB_V2_ACUM = vWeb * cambioSigo;
                //        }
                //    }

                //    #endregion
                //}

                celda = GetMes();

                if (concepto.Id == 41) //41_Efectivo e Inversiones
                    celda += "8";
                else if (concepto.Id == 42) //42_Documentos por Cobrar Unidades
                    celda += "10";
                else if (concepto.Id == 43) //43_Cuentas por Cobrar Unidades
                    celda += "11";
                else if (concepto.Id == 44) //44_Cuentas por Cobrar Refacciones
                    celda += "12";
                else if (concepto.Id == 45) //45_Cuentas por Cobrar Servicio
                    celda += "13";
                else if (concepto.Id == 46) //46_Reserva Incobrables
                    celda += "14";
                else if (concepto.Id == 47) //47_Planta Activo
                    celda += "17";
                else if (concepto.Id == 48) //48_Deudores Diversos y Funcionarios y Empleados
                    celda += "18";
                else if (concepto.Id == 49) //49_Partes Relacionadas Activo
                    celda += "19";
                else if (concepto.Id == 50) //50_Impuestos por Recuperar
                    celda += "22";
                else if (concepto.Id == 51) //51_Pagos Anticipados
                    celda += "23";
                else if (concepto.Id == 52) //52_Inventario Autos Nuevos
                    celda += "25";
                else if (concepto.Id == 53) //53_Inventario Autos Seminuevos
                    celda += "26";
                else if (concepto.Id == 54) //54_Inventario Refacciones y Accesorios
                    celda += "27";
                else if (concepto.Id == 55) //55_Inventarios Otros y Proceso
                    celda += "28";
                //else if (concepto.Id == 10617) //10617	Materiales Obsoletos
                //    celda += "";
                else if (concepto.Id == 58) //58_Terreno y Edificio
                    celda += "36";
                else if (concepto.Id == 59) //59_Mobiliario y Equipo
                    celda += "37";
                else if (concepto.Id == 60) //60_Equipo de Transporte
                    celda += "38";
                else if (concepto.Id == 61) //61_Depreciación Acumulada
                    celda += "39";
                else if (concepto.Id == 62) //62_Otros Activos
                    celda += "42";
                //else if (concepto.Id == 10609) //10609_Activo por derecho de uso
                //    celda += "";
                else if (concepto.Id == 63) //63_Mejoras Inmuebles Arrendados
                    celda += "43";
                else if (concepto.Id == 64) //64_Neto de Actualizaciones
                    celda += "44";
                else if (concepto.Id == 56) //56_Inversiones Permanentes en Acciones
                    celda += "33";
                else if (concepto.Id == 66) //66_Planta Pasivo
                    celda += "52";
                else if (concepto.Id == 67) //67_Impuestos por Pagar
                    celda += "53";
                else if (concepto.Id == 68) //68_Anticipos de Clientes
                    celda += "54";
                else if (concepto.Id == 70) //70_Proveedores
                    celda += "56";
                else if (concepto.Id == 71) //71_Acreedores Diversos
                    celda += "57";
                else if (concepto.Id == 72) //72_P.T.U.
                    celda += "58";
                else if (concepto.Id == 73) //73_Partes Relacionadas Pasivo
                    celda += "59";
                else if (concepto.Id == 74) //74_Documentos por Pagar
                    celda += "60";
                else if (concepto.Id == 75) //75_Otras Provisiones y Cuentas por Pagar
                    celda += "61";
                else if (concepto.Id == 76) //76_I.S.R. Diferido
                    celda += "62";
                else if (concepto.Id == 79) //79_Intereses
                    celda += "67";
                else if (concepto.Id == 80) //80_Prima de Antigüedad
                    celda += "68";
                //else if (concepto.Id == 10610) //10610_Pasivo acumulado (Arrendamiento)
                //    celda += "";
                else if (concepto.Id == 81) //81_Capital Social
                    celda += "76";
                else if (concepto.Id == 82) //82_Reserva Legal
                    celda += "77";
                else if (concepto.Id == 83) //83_Resultado de Ejercicios Anteriores
                    celda += "78";
                //else if (concepto.Id == 84) //84_Resultado del Ejercicio
                //    celda += "79";
                else if (concepto.Id == 85) //85_Actualización del Capital Contable
                    celda += "80";
                //else if (concepto.Id == 10612) //10612_Aportación para futuro aumentos de capital
                //    celda += "";

                if (celda.Length >= 2)
                {
                    if (version == "V1" || version == "V2")
                        ws.get_Range(celda, celda).Formula = rep.WEB_V1;
                    else
                        ws.get_Range(celda, celda).Formula = rep.WEB_V2;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + concepto.Id + "_" + concepto.NombreConcepto);                
            }
        }

        public void LlenaPestaniaBGV2(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<ReporteBalanceGeneral> lstBG = ReporteBalanceGeneral.ListarGeneral(_db, idAgencia, anio);
            List<ReporteBalanceGeneral> lstBGV2 = ReporteBalanceGeneral.ListarGeneralExtralibros(_db, idAgencia, anio);

            ReporteBalanceGeneral resEjeV1 = lstBG.Find(x => x.IdConcepto == 84);
            ReporteBalanceGeneral resEjeV2 = lstBGV2.Find(x => x.IdConcepto == 84);

            decimal diffV1vsV2 = 0;

            if (mes == 1)
                diffV1vsV2 = resEjeV1.Eenero - resEjeV2.Eenero;
            else if (mes == 2)
                diffV1vsV2 = resEjeV1.Febrero - resEjeV2.Febrero;
            else if (mes == 3)
                diffV1vsV2 = resEjeV1.Marzo - resEjeV2.Marzo;
            else if (mes == 4)
                diffV1vsV2 = resEjeV1.Abril - resEjeV2.Abril;
            else if (mes == 5)
                diffV1vsV2 = resEjeV1.Mayo - resEjeV2.Mayo;
            else if (mes == 6)
                diffV1vsV2 = resEjeV1.Junio - resEjeV2.Junio;
            else if (mes == 7)
                diffV1vsV2 = resEjeV1.Julio - resEjeV2.Julio;
            else if (mes == 8)
                diffV1vsV2 = resEjeV1.Agosto - resEjeV2.Agosto;
            else if (mes == 9)
                diffV1vsV2 = resEjeV1.Septiembre - resEjeV2.Septiembre;
            else if (mes == 10)
                diffV1vsV2 = resEjeV1.Octubre - resEjeV2.Octubre;
            else if (mes == 11)
                diffV1vsV2 = resEjeV1.Noviembre - resEjeV2.Noviembre;
            else if (mes == 12)
                diffV1vsV2 = resEjeV1.Diciembre - resEjeV2.Diciembre;

            celdaMes = GetMes();

            foreach (ReporteBalanceGeneral bg in lstBG)
            {
                if (bg.IdConcepto == 41) //41_Efectivo e Inversiones
                    celda = celdaMes + "8";
                else if (bg.IdConcepto == 42) //42_Documentos por Cobrar Unidades
                    celda = celdaMes + "10";
                else if (bg.IdConcepto == 43) //43_Cuentas por Cobrar Unidades
                    celda = celdaMes + "11";
                else if (bg.IdConcepto == 44) //44_Cuentas por Cobrar Refacciones
                    celda = celdaMes + "12";
                else if (bg.IdConcepto == 45) //45_Cuentas por Cobrar Servicio
                    celda = celdaMes + "13";
                else if (bg.IdConcepto == 46) //46_Reserva Incobrables
                    celda = celdaMes + "14";
                else if (bg.IdConcepto == 47) //47_Planta Activo
                    celda = celdaMes + "17";
                else if (bg.IdConcepto == 48) //48_Deudores Diversos y Funcionarios y Empleados
                    celda = celdaMes + "18";
                else if (bg.IdConcepto == 49) //49_Partes Relacionadas Activo
                    celda = celdaMes + "19";
                else if (bg.IdConcepto == 50) //50_Impuestos por Recuperar
                    celda = celdaMes + "22";
                else if (bg.IdConcepto == 51) //51_Pagos Anticipados
                    celda = celdaMes + "23";
                else if (bg.IdConcepto == 52) //52_Inventario Autos Nuevos
                    celda = celdaMes + "25";
                else if (bg.IdConcepto == 53) //53_Inventario Autos Seminuevos
                    celda = celdaMes + "26";
                else if (bg.IdConcepto == 54) //54_Inventario Refacciones y Accesorios
                    celda = celdaMes + "27";
                else if (bg.IdConcepto == 55) //55_Inventarios Otros y Proceso
                    celda = celdaMes + "28";
                //else if (concepto.IdConcepto == 10617) //10617	Materiales Obsoletos
                //    celda = celdaMes + "";
                else if (bg.IdConcepto == 58) //58_Terreno y Edificio
                    celda = celdaMes + "36";
                else if (bg.IdConcepto == 59) //59_Mobiliario y Equipo
                    celda = celdaMes + "37";
                else if (bg.IdConcepto == 60) //60_Equipo de Transporte
                    celda = celdaMes + "38";
                else if (bg.IdConcepto == 61) //61_Depreciación Acumulada
                    celda = celdaMes + "39";
                else if (bg.IdConcepto == 62) //62_Otros Activos
                    celda = celdaMes + "42";
                //else if (concepto.IdConcepto == 10609) //10609_Activo por derecho de uso
                //    celda = celdaMes + "";
                else if (bg.IdConcepto == 63) //63_Mejoras Inmuebles Arrendados
                    celda = celdaMes + "43";
                else if (bg.IdConcepto == 64) //64_Neto de Actualizaciones
                    celda = celdaMes + "44";
                else if (bg.IdConcepto == 56) //56_Inversiones Permanentes en Acciones
                    celda = celdaMes + "33";
                else if (bg.IdConcepto == 66) //66_Planta Pasivo
                    celda = celdaMes + "52";
                else if (bg.IdConcepto == 67) //67_Impuestos por Pagar
                    celda = celdaMes + "53";
                else if (bg.IdConcepto == 68) //68_Anticipos de Clientes
                    celda = celdaMes + "54";
                else if (bg.IdConcepto == 70) //70_Proveedores
                    celda = celdaMes + "56";
                else if (bg.IdConcepto == 71) //71_Acreedores Diversos
                    celda = celdaMes + "57";
                else if (bg.IdConcepto == 72) //72_P.T.U.
                    celda = celdaMes + "58";
                else if (bg.IdConcepto == 73) //73_Partes Relacionadas Pasivo
                    celda = celdaMes + "59";
                else if (bg.IdConcepto == 74) //74_Documentos por Pagar
                    celda = celdaMes + "60";
                else if (bg.IdConcepto == 75) //75_Otras Provisiones y Cuentas por Pagar
                    celda = celdaMes + "61";
                else if (bg.IdConcepto == 76) //76_I.S.R. Diferido
                    celda = celdaMes + "62";
                else if (bg.IdConcepto == 79) //79_Intereses
                    celda = celdaMes + "67";
                else if (bg.IdConcepto == 80) //80_Prima de Antigüedad
                    celda = celdaMes + "68";
                //else if (concepto.Id == 10610) //10610_Pasivo acumulado (Arrendamiento)
                //    celda = celdaMes + "";
                else if (bg.IdConcepto == 81) //81_Capital Social
                    celda = celdaMes + "76";
                else if (bg.IdConcepto == 82) //82_Reserva Legal
                    celda = celdaMes + "77";
                else if (bg.IdConcepto == 83) //83_Resultado de Ejercicios Anteriores
                    celda = celdaMes + "78";
                //else if (concepto.Id == 84) //84_Resultado del Ejercicio
                //    celda = celdaMes + "79";
                else if (bg.IdConcepto == 85) //85_Actualización del Capital Contable
                    celda = celdaMes + "80";
                //else if (concepto.Id == 10612) //10612_Aportación para futuro aumentos de capital
                //    celda = celdaMes + "";

                if (version == "V1")
                {
                    if (celda.Length >= 2)
                    {
                        if (mes == 1)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Eenero / 1000);
                        else if (mes == 2)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Febrero / 1000);
                        else if (mes == 3)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Marzo / 1000);
                        else if (mes == 4)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Abril / 1000);
                        else if (mes == 5)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Mayo / 1000);
                        else if (mes == 6)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Junio / 1000);
                        else if (mes == 7)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Julio / 1000);
                        else if (mes == 8)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Agosto / 1000);
                        else if (mes == 9)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Septiembre / 1000);
                        else if (mes == 10)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Octubre / 1000);
                        else if (mes == 11)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Noviembre / 1000);
                        else if (mes == 12)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(bg.Diciembre / 1000);
                    }
                }
                else
                {
                    if (celda.Length >= 2)
                    {
                        if (mes == 1)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Eenero + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Eenero / 1000);
                        else if (mes == 2)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Febrero + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Febrero / 1000);
                        else if (mes == 3)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Marzo + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Marzo / 1000);
                        else if (mes == 4)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Abril + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Abril / 1000);
                        else if (mes == 5)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Mayo + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Mayo / 1000);
                        else if (mes == 6)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Junio + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Junio / 1000);
                        else if (mes == 7)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Julio + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Julio / 1000);
                        else if (mes == 8)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Agosto + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Agosto / 1000);
                        else if (mes == 9)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Septiembre + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Septiembre / 1000);
                        else if (mes == 10)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Octubre + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Octubre / 1000);
                        else if (mes == 11)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Noviembre + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Noviembre / 1000);
                        else if (mes == 12)
                            ws.get_Range(celda, celda).Formula = bg.IdConcepto == 75 ? Convert.ToInt32((bg.Diciembre + diffV1vsV2) / 1000) : Convert.ToInt32(bg.Diciembre / 1000);
                    }
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + bg.IdConcepto + "_" + bg.NombreConcepto);
            }
        }

        public void LlenaPestaniaCuentas(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";
            decimal importe = 0;
            bool banNegativo = false;

            List<Cuentas> lstCuentas = Cuentas.Listar(_db, idAgencia, anio);

            celdaMes = GetMes();

            foreach (Cuentas cuenta in lstCuentas)
            {
                if (cuenta.IdConcepto == 709) //709_Cuentas por Cobrar Unidades Nuevas (Bancos, cartera 1130)
                    celda = celdaMes + "3";
                else if (cuenta.IdConcepto == 710) //710_Cuentas por Cobrar Financieras de Marca
                    celda = celdaMes + "4";
                else if (cuenta.IdConcepto == 711 && cuenta.Concepto.Contains("Cuentas por Cobrar Unidades Nuevas")) //711_Cuentas por Cobrar Unidades Nuevas
                    celda = celdaMes + "7";
                else if (cuenta.IdConcepto == 712) //712_Cuentas por Cobrar Garantias
                    celda = celdaMes + "9";
                else if (cuenta.IdConcepto == 713) //713_Venta Mano de Obra Taller
                    celda = celdaMes + "12";
                else if (cuenta.IdConcepto == 714) //714_Devolución Venta Mano de Obra Taller
                    celda = celdaMes + "13";
                else if (cuenta.IdConcepto == 715) //715_Costo de Venta Mano de Obra Taller
                    celda = celdaMes + "14";
                else if (cuenta.IdConcepto == 716) //716_Recuperacion de Mano de Obra
                    celda = celdaMes + "15";
                else if (cuenta.IdConcepto == 717) //717_Costo M.O. Mecanica Proceso
                    celda = celdaMes + "16";
                else if (cuenta.IdConcepto == 718 && cuenta.Concepto.Contains("Venta Mano de Obra Garantias")) //718_Venta Mano de Obra Garantias
                    celda = celdaMes + "19";
                else if (cuenta.IdConcepto == 719) //719_Devolución Venta Mano de Obra Garantias
                    celda = celdaMes + "20";
                else if (cuenta.IdConcepto == 720) //720_Costo de Venta Mano de Obra Garantias
                    celda = celdaMes + "21";
                else if (cuenta.IdConcepto == 721 && cuenta.Concepto.Contains("Ventas T.O.T.")) //721_Ventas T.O.T.
                    celda = celdaMes + "24";
                else if (cuenta.IdConcepto == 722) //722_Devolución T.O.T.
                    celda = celdaMes + "25";
                else if (cuenta.IdConcepto == 723) //723_Costo de Vetas T.O.T.
                    celda = celdaMes + "26";
                else if (cuenta.IdConcepto == 724 && cuenta.Concepto.Contains("Venta Materiales Diversos.")) //724_Venta Materiales Diversos.
                    celda = celdaMes + "29";
                else if (cuenta.IdConcepto == 725) //725_Devolución Venta Materiales Diversos.
                    celda = celdaMes + "30";
                else if (cuenta.IdConcepto == 726) //726_Costo de Venta Materiales Diversos.
                    celda = celdaMes + "31";
                else if (cuenta.IdConcepto == 727 && cuenta.Concepto.Contains("Venta Refacciones Mayoreo.")) //727_Venta Refacciones Mayoreo.
                    celda = celdaMes + "35";
                else if (cuenta.IdConcepto == 728) //728_Devolución Venta Refacciones Mayoreo.
                    celda = celdaMes + "36";
                else if (cuenta.IdConcepto == 729) //729_Costo de Venta Refacciones Mayoreo.
                    celda = celdaMes + "37";
                else if (cuenta.IdConcepto == 730) //730_Ingresos por Bonificaciones Refacciones
                    celda = celdaMes + "38";
                else if (cuenta.IdConcepto == 731 && cuenta.Concepto.Contains("Venta Refacciones Mostrador.")) //731_Venta Refacciones Mostrador.
                    celda = celdaMes + "41";
                else if (cuenta.IdConcepto == 732) //732_Devolución Venta Refacciones Mostrador.
                    celda = celdaMes + "42";
                else if (cuenta.IdConcepto == 733) //733_Costo de Venta Refacciones Mostrador.
                    celda = celdaMes + "43";
                else if (cuenta.IdConcepto == 734 && cuenta.Concepto.Contains("Venta Refacciones Taller.")) //734_Venta Refacciones Taller.
                    celda = celdaMes + "46";
                else if (cuenta.IdConcepto == 735) //735_Devolución Venta Refacciones Taller.
                    celda = celdaMes + "47";
                else if (cuenta.IdConcepto == 736) //736_Costo de Venta Refacciones Taller.
                    celda = celdaMes + "48";
                else if (cuenta.IdConcepto == 737 && cuenta.Concepto.Contains("Venta Refacciones Cía. de Seguro.")) //737_Venta Refacciones Cía. de Seguro.
                    celda = celdaMes + "51";
                else if (cuenta.IdConcepto == 738) //738_Devolución Venta Refacciones Cía. de Seguro.
                    celda = celdaMes + "52";
                else if (cuenta.IdConcepto == 739) //739_Costo de Venta Refacciones Cía. de Seguro.
                    celda = celdaMes + "53";
                else if (cuenta.IdConcepto == 740 && cuenta.Concepto.Contains("Venta Refacciones Garantía.")) //740_Venta Refacciones Garantía.
                    celda = celdaMes + "56";
                else if (cuenta.IdConcepto == 741) //741_Devolución Venta Refacciones Garantía.
                    celda = celdaMes + "57";
                else if (cuenta.IdConcepto == 742) //742_Costo de Venta Refacciones Garantía.
                    celda = celdaMes + "58";
                else if (cuenta.IdConcepto == 743 && cuenta.Concepto.Contains("Venta Refacciones Otras Mercancias.")) //743_Venta Refacciones Otras Mercancias.
                    celda = celdaMes + "61";
                else if (cuenta.IdConcepto == 744) //744_Devolución Venta Refacciones Otras Mercancias.
                    celda = celdaMes + "62";
                else if (cuenta.IdConcepto == 745) //745_Costo de Venta Refacciones Otras Mercancias.
                    celda = celdaMes + "63";
                else if (cuenta.IdConcepto == 746 && cuenta.Concepto.Contains("Venta Refacciones Accesorios.")) //746_Venta Refacciones Accesorios.
                    celda = celdaMes + "66";
                else if (cuenta.IdConcepto == 747) //747_Devolución Venta Refacciones Accesorios.
                    celda = celdaMes + "67";
                else if (cuenta.IdConcepto == 748) //748_Costo de Venta Refacciones Accesorios.
                    celda = celdaMes + "68";
                else if (cuenta.IdConcepto == 749 && cuenta.Concepto.Contains("Venta Hojalateria y Pintura")) //749_Venta Hojalateria y Pintura
                    celda = celdaMes + "72";
                else if (cuenta.IdConcepto == 750) //750_Devolución Venta Hojalateria y Pintura
                    celda = celdaMes + "73";
                else if (cuenta.IdConcepto == 751) //751_Costo de Venta Hojalateria y Pintura
                    celda = celdaMes + "74";
                else if (cuenta.IdConcepto == 752) //752_Costo de Venta Mano de Obra Hojalateria y Pintura Proceso
                    celda = celdaMes + "75";
                else if (cuenta.IdConcepto == 753) //753_Recuperación Mano de Obra Hojalateria y Pintura
                    celda = celdaMes + "76";
                else if (cuenta.IdConcepto == 754 && cuenta.Concepto.Contains("Venta Materiales Diversos Hojalateria y Pintura")) //754_Venta Materiales Diversos Hojalateria y Pintura
                    celda = celdaMes + "79";
                else if (cuenta.IdConcepto == 755) //755_Devolución Venta Materiales Diversos Hojalateria y Pintura
                    celda = celdaMes + "80";
                else if (cuenta.IdConcepto == 762) //762_RECUPERACIÓN MANO DE OBRA HOJALATERIA Y PINTURA MAT DIVERSOS
                    celda = celdaMes + "81";

                if (cuenta.IdConcepto == 714 || cuenta.IdConcepto == 715 || cuenta.IdConcepto == 717 || cuenta.IdConcepto == 719 || cuenta.IdConcepto == 720 ||
                    cuenta.IdConcepto == 722 || cuenta.IdConcepto == 723 || cuenta.IdConcepto == 725 || cuenta.IdConcepto == 726 || cuenta.IdConcepto == 728 ||
                    cuenta.IdConcepto == 729 || cuenta.IdConcepto == 732 || cuenta.IdConcepto == 733 || cuenta.IdConcepto == 735 || cuenta.IdConcepto == 736 ||
                    cuenta.IdConcepto == 738 || cuenta.IdConcepto == 739 || cuenta.IdConcepto == 741 || cuenta.IdConcepto == 742 || cuenta.IdConcepto == 744 ||
                    cuenta.IdConcepto == 745 || cuenta.IdConcepto == 747 || cuenta.IdConcepto == 748 || cuenta.IdConcepto == 750 || cuenta.IdConcepto == 751 ||
                    cuenta.IdConcepto == 752 || cuenta.IdConcepto == 755)
                    banNegativo = true;

                if (celda.Length >= 2)
                {
                    if (mes == 1)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Ene / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Ene / 1000);
                    }
                    else if (mes == 2)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Feb / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Feb / 1000);
                    }
                    else if (mes == 3)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Mar / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Mar / 1000);
                    }
                    else if (mes == 4)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Abr / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Abr / 1000);
                    }
                    else if (mes == 5)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.May / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.May / 1000);
                    }
                    else if (mes == 6)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Jun / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Jun / 1000);
                    }
                    else if (mes == 7)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Jul / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Jul / 1000);
                    }
                    else if (mes == 8)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Ago / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Ago / 1000);
                    }
                    else if (mes == 9)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Sep / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Sep / 1000);
                    }
                    else if (mes == 10)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Oct / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Oct / 1000);
                    }
                    else if (mes == 11)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Nov / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Nov / 1000);
                    }
                    else if (mes == 12)
                    {
                        if (banNegativo)
                            ws.get_Range(celda, celda).Formula = importe = (Convert.ToInt32(cuenta.Dic / 1000) * -1);
                        else
                            ws.get_Range(celda, celda).Formula = importe = Convert.ToInt32(cuenta.Dic / 1000);
                    }
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + cuenta.IdConcepto + "_" + cuenta.Concepto + " = " + importe);

                importe = 0;
                banNegativo = false;
            }
        }

        public void LlenaPestaniaVP(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";
            bool banEsPorcentaje = false;

            List<VolumenYPorcentaje> lstVP;

            if (version == "V1")
                lstVP = VolumenYPorcentaje.Listar(_db, idAgencia, anio);
            else
                lstVP = VolumenYPorcentaje.ListarExtralibros(_db, idAgencia, anio);

            celdaMes = GetMes();

            foreach (VolumenYPorcentaje vp in lstVP)
            {
                if (vp.IdConcept == 89) //89_UTILIDAD BRUTA MENUDEO
                    celda = celdaMes + "8";
                else if (vp.IdConcept == 90) //90_UTILIDAD BRUTA PROMEDIO UNIDAD NUEVA VENDIDA
                    celda = celdaMes + "9";
                else if (vp.IdConcept == 92) //92_UTILIDAD BRUTA FLOTILLAS
                    celda = celdaMes + "12";
                else if (vp.IdConcept == 93) //93_U. B. PROMEDIO UNIDAD NUEVA VENDIDA FLOTILLA
                    celda = celdaMes + "13";
                else if (vp.IdConcept == 95) //95_UTILIDAD BRUTA INTERNET
                    celda = celdaMes + "16";
                else if (vp.IdConcept == 96) //96_U. B. PROMEDIO UNIDAD NUEVA VENDIDA INTERNET
                    celda = celdaMes + "17";
                else if (vp.IdConcept == 98) //96_U. B. PROMEDIO UNIDAD NUEVA VENDIDA INTERNET
                    celda = celdaMes + "20";

                if (vp.IdConcept == 89 || vp.IdConcept == 92 || vp.IdConcept == 95 || vp.IdConcept == 98)
                    banEsPorcentaje = true;

                if (celda.Length >= 2)
                {
                    if (mes == 1)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Ene / 100) : vp.Ene;
                    else if (mes == 2)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Feb / 100) : vp.Feb;
                    else if (mes == 3)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Mar / 100) : vp.Mar;
                    else if (mes == 4)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Abr / 100) : vp.Abr;
                    else if (mes == 5)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.May / 100) : vp.May;
                    else if (mes == 6)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Jun / 100) : vp.Jun;
                    else if (mes == 7)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Jul / 100) : vp.Jul;
                    else if (mes == 8)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Ago / 100) : vp.Ago;
                    else if (mes == 9)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Sep / 100) : vp.Sep;
                    else if (mes == 10)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Oct / 100) : vp.Oct;
                    else if (mes == 11)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Nov / 100) : vp.Nov;
                    else if (mes == 12)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (vp.Dic / 100) : vp.Dic;
                }

                celda = "";
                banEsPorcentaje = false;

                Console.WriteLine("[CONCEPTO]: " + vp.IdConcept + "_" + vp.Concepto);
            }
        }

        public void LlenaPestaniaAG(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<AspectoGeneral> lstAG = AspectoGeneral.ListarExtralibros(_db, idAgencia, anio);
            List<ReportePersonalAG> lstAGPersonal = AspectoGeneral.ListarPersonal(_db, idAgencia, anio);

            celdaMes = GetMes();

            foreach (AspectoGeneral ag in lstAG)
            {
                if (ag.IdConcepto == 670) //670_VENTAS FIDEICOMISO MES
                    celda = celdaMes + "6";

                if (celda.Length >= 2)
                {
                    if (mes == 1)
                        ws.get_Range(celda, celda).Formula = ag.Ene / 1000;
                    else if (mes == 2)
                        ws.get_Range(celda, celda).Formula = ag.Feb / 1000;
                    else if (mes == 3)
                        ws.get_Range(celda, celda).Formula = ag.Mar / 1000;
                    else if (mes == 4)
                        ws.get_Range(celda, celda).Formula = ag.Abr / 1000;
                    else if (mes == 5)
                        ws.get_Range(celda, celda).Formula = ag.May / 1000;
                    else if (mes == 6)
                        ws.get_Range(celda, celda).Formula = ag.Jun / 1000;
                    else if (mes == 7)
                        ws.get_Range(celda, celda).Formula = ag.Jul / 1000;
                    else if (mes == 8)
                        ws.get_Range(celda, celda).Formula = ag.Ago / 1000;
                    else if (mes == 9)
                        ws.get_Range(celda, celda).Formula = ag.Sep / 1000;
                    else if (mes == 10)
                        ws.get_Range(celda, celda).Formula = ag.Oct / 1000;
                    else if (mes == 11)
                        ws.get_Range(celda, celda).Formula = ag.Nov / 1000;
                    else if (mes == 12)
                        ws.get_Range(celda, celda).Formula = ag.Dic / 1000;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + ag.IdConcepto + "_" + ag.Concepto);
            }

            foreach (ReportePersonalAG rpag in lstAGPersonal)
            {
                if (rpag.FIFNTIPEMP == 1) //1_PERSONAL EN NOMINA
                {
                    if (rpag.FIFNIDDEP == 1) //1_AUTOS NUEVOS
                        celda = celdaMes + "38";
                    else if (rpag.FIFNIDDEP == 4) //4_AUTOS SEMINUEVOS
                        celda = celdaMes + "39";
                    else if (rpag.FIFNIDDEP == 5) //5_SERVICIO
                        celda = celdaMes + "40";
                    else if (rpag.FIFNIDDEP == 6) //6_HOJALATERIA Y PINTURA
                        celda = celdaMes + "41";
                    else if (rpag.FIFNIDDEP == 8) //8_REFACCIONES
                        celda = celdaMes + "42";
                    else if (rpag.FIFNIDDEP == 10) //10_ADMINISTRACION
                        celda = celdaMes + "43";
                    else if (rpag.FIFNIDDEP == 55) //55_VIGILANCIA Y ASEO
                        celda = celdaMes + "44";
                }
                else
                {
                    if (rpag.FIFNIDDEP == 1) //1_AUTOS NUEVOS
                        celda = celdaMes + "47";
                    else if (rpag.FIFNIDDEP == 4) //4_AUTOS SEMINUEVOS
                        celda = celdaMes + "48";
                    else if (rpag.FIFNIDDEP == 5) //5_SERVICIO
                        celda = celdaMes + "49";
                    else if (rpag.FIFNIDDEP == 6) //6_HOJALATERIA Y PINTURA
                        celda = celdaMes + "50";
                    else if (rpag.FIFNIDDEP == 8) //8_REFACCIONES
                        celda = celdaMes + "51";
                    else if (rpag.FIFNIDDEP == 10) //10_ADMINISTRACION
                        celda = celdaMes + "52";
                    else if (rpag.FIFNIDDEP == 55) //55_VIGILANCIA Y ASEO
                        celda = celdaMes + "53";
                }

                if (celda.Length >= 2)
                {
                    if (mes == 1)
                        ws.get_Range(celda, celda).Formula = rpag.ENE;
                    else if (mes == 2)
                        ws.get_Range(celda, celda).Formula = rpag.FEB;
                    else if (mes == 3)
                        ws.get_Range(celda, celda).Formula = rpag.MAR;
                    else if (mes == 4)
                        ws.get_Range(celda, celda).Formula = rpag.ABR;
                    else if (mes == 5)
                        ws.get_Range(celda, celda).Formula = rpag.MAY;
                    else if (mes == 6)
                        ws.get_Range(celda, celda).Formula = rpag.JUN;
                    else if (mes == 7)
                        ws.get_Range(celda, celda).Formula = rpag.JUL;
                    else if (mes == 8)
                        ws.get_Range(celda, celda).Formula = rpag.AGO;
                    else if (mes == 9)
                        ws.get_Range(celda, celda).Formula = rpag.SEP;
                    else if (mes == 10)
                        ws.get_Range(celda, celda).Formula = rpag.OCT;
                    else if (mes == 11)
                        ws.get_Range(celda, celda).Formula = rpag.NOV;
                    else if (mes == 12)
                        ws.get_Range(celda, celda).Formula = rpag.DIC;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + rpag.FIFNTIPEMP + "_" + rpag.FIFNIDDEP);
            }
        }

        public void LlenaPestaniaCC(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";
            string idsAgencias = "";
            int r = 0;

            foreach (int i in GetAgenciasYSucursales(_db, idAgencia))
            {
                idsAgencias += i + ",";
            }

            idsAgencias = idsAgencias.Remove(idsAgencias.Length - 1, 1);
                        
            celdaMes = GetMes();

            r = Convert.ToInt32(celdaMes);

            PeriodoContable periodo = PeriodoContable.BuscarPorMesAnio(_db, mes, anio);

            query = "SELECT \r\n" +
                "FACTU . FICAIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, FACTU . FICAIDFACT ID_FACTURA, FACTU . FICAFOLIN FOLIO_FACTURA, " +
                "FACTU . FFCAFECHA FECHA_FACTURA, FACTU . FICAIDCLIE ID_PERSONA, \r\n" +
                "CASE WHEN PERMO . FDPMIDPERS IS NOT NULL THEN CONCAT(CONCAT(PERSO . FDPEIDPERS, '_'), TRIM(PERMO . FSPMRAZON))\r\n" +
                "ELSE CONCAT(CONCAT(CONCAT(CONCAT(CONCAT(CONCAT(PERSO . FDPEIDPERS, '_'), TRIM(PERFI . FSPFNOMBRE)), ' '), TRIM(PERFI . FSPFAPATER)), ' '), TRIM(PERFI . FSPFAMATER)) END PERSONA, \r\n" +
                "FACTU . FFCAFECHAC FECHA_CANCELACION,\r\nFACAN.FSCANUMSER NUMERO_DE_SERIE,\r\n" +
                "SALID.FIAUFOLSAL FOLIO_DE_SALIDA, SALID.FFAUFESALI FECHA_DE_SALIDA, FACAN . FICAIDTIVT ID_TIPO_VENTA, FACAN . FSCATIPVTA TIPO_VENTA\r\n" +
                "FROM [PREFIX]CAJA . CAEFACTU FACTU\r\n" +
                "INNER JOIN [PREFIX]CAJA . CAEFACAN FACAN ON FACTU . FICAIDCIAU = FACAN . FICAIDCIAU AND FACTU . FICAIDFACT = FACAN . FICAIDFACT\r\n" +
                "INNER JOIN [PREFIX]GRAL . GECCIAUN CIAUN ON FACTU . FICAIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN [PREFIX]PERS . CTEPERSO PERSO ON FACTU . FICAIDCLIE = PERSO . FDPEIDPERS\r\n" +
                "LEFT JOIN [PREFIX]PERS . CTDPERMO PERMO ON PERSO . FDPEIDPERS = PERMO . FDPMIDPERS AND PERMO . FDPMESTATU = 1\r\n" +
                "LEFT JOIN [PREFIX]PERS . CTCPERFI PERFI ON PERSO . FDPEIDPERS = PERFI . FDPFIDPERS AND PERFI . FDPFESTATU = 1\r\n" +
                "LEFT JOIN [PREFIX]AUSA . AUCSALID SALID ON FACTU . FICAIDCIAU = SALID . FIAUIDCIAU AND FACTU . FICAIDFACT = SALID . FIAUIDFACT AND SALID . FIAUIDTPSL = 9 " +
                "AND SALID . FIAUSTATUS = 1  \r\n" +
                "WHERE \r\n" +
                "FACAN . FICAIDCIAU IN (" + idsAgencias + ")\r\n" +
                "AND FACTU . FFCAFECHA >= '" + periodo.FechaInicial.ToString("yyyy-MM-dd") + "' " +
                "AND FACTU . FFCAFECHA <= '" + periodo.FechaFinal.ToString("yyyy-MM-dd") + "'\r\n" +
                "AND FACAN . FICAIDTIVT IN (41,42,43,44,47,92)\r\n" +
                "ORDER BY FACAN.FSCANUMSER, FACTU . FICAFOLIN";

            //TIPOS DE VENTA RELACIONADOS A INTERCOMPAÑIAS E INTERCAMBIOS
            //41  INTERCAMBIOS
            //42  INTERCAMBIOS PTO VTA
            //43  INTERCOMPAÑIAS
            //44  INTERCOMPAÑIAS PTO VTA
            //47  OTROS CONCESIONARIOS
            //92  SESION DE PASIVO

            System.Data.DataTable dtVentasInter = _db.GetDataTable(query);

            celda = "N" + celdaMes;

            ws.get_Range(celda, celda).Formula = dtVentasInter.Rows.Count;
            Console.WriteLine("VENTAS INTERCAMBIOS: " + dtVentasInter.Rows.Count);
        }

        public void LlenaPestaniaAF(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";
            bool banEsPorcentaje = false;

            List<AspectoFinanciero> lstAF = AspectoFinanciero.Listar(_db, idAgencia, anio);

            celdaMes = GetMes();

            foreach (AspectoFinanciero af in lstAF)
            {
                if (af.IdConcepto == 117) //117_UNIDADES NUEVAS SIN INTERESES
                    celda = celdaMes + "7";
                else if (af.IdConcepto == 118) //118_UNIDADES NUEVAS CON INTERESES
                    celda = celdaMes + "8";
                else if (af.IdConcepto == 119) //119_TASA APLICABLE UN
                    celda = celdaMes + "9";
                else if (af.IdConcepto == 10972) //10972_INTERESES PAGADOS UN
                    celda = celdaMes + "10";
                else if (af.IdConcepto == 10415) //10415_UNIDADES SEMINUEVAS SIN INTERESES
                    celda = celdaMes + "11";
                else if (af.IdConcepto == 10414) //10414_UNIDADES SEMINUEVAS CON INTERESES
                    celda = celdaMes + "12";
                else if (af.IdConcepto == 10416) //10416_TASA APLICABLE USN
                    celda = celdaMes + "13";
                else if (af.IdConcepto == 10973) //10973_INTERESES PAGADOS USN
                    celda = celdaMes + "14";
                else if (af.IdConcepto == 10417) //10417_REFACCIONES SIN INTERESES
                    celda = celdaMes + "15";
                else if (af.IdConcepto == 120) //120_REFACCIONES CON INTERESES
                    celda = celdaMes + "16";
                else if (af.IdConcepto == 10418) //10418_TASA APLICABLE REFA
                    celda = celdaMes + "17";
                else if (af.IdConcepto == 121) //121_INTERESES PAGADOS REFA
                    celda = celdaMes + "18";
                //else if (af.IdConcepto == 123) //123_OTROS
                //    celda = celdaMes + "28";
                else if (af.IdConcepto == 122) //122_ORGANIZACION HERPA
                    celda = celdaMes + "25";
                else if (af.IdConcepto == 10974) //10974_ORGANIZACIÓN HERPA TASA
                    celda = celdaMes + "26";
                else if (af.IdConcepto == 10975) //10975_ORGANIZACIÓN HERPA INTERESES PAGADOS
                    celda = celdaMes + "27";
                else if (af.IdConcepto == 10976) //10976_ACTINVER
                    celda = celdaMes + "28";
                else if (af.IdConcepto == 10977) //10977_ACTINVER TASA
                    celda = celdaMes + "29";
                else if (af.IdConcepto == 10978) //10978_ACTINVER INTERESES PAGADOS
                    celda = celdaMes + "30";
                else if (af.IdConcepto == 10979) //10979_BANBAJIO
                    celda = celdaMes + "31";
                else if (af.IdConcepto == 10980) //10980_BANBAJIO TASA
                    celda = celdaMes + "32";
                else if (af.IdConcepto == 10981) //10981_BANBAJIO INTERESES PAGADOS
                    celda = celdaMes + "33";
                else if (af.IdConcepto == 10982) //10982_BAM
                    celda = celdaMes + "34";
                else if (af.IdConcepto == 10983) //10983_BAM TASA
                    celda = celdaMes + "35";
                else if (af.IdConcepto == 10984) //10984_BAM INTERESES PAGADOS
                    celda = celdaMes + "36";
                else if (af.IdConcepto == 10985) //10985_BANCOMEXT
                    celda = celdaMes + "37";
                else if (af.IdConcepto == 10986) //10986_BANCOMEXT TASA
                    celda = celdaMes + "38";
                else if (af.IdConcepto == 10987) //10987_BANCOMEXT INTERESES PAGADOS
                    celda = celdaMes + "39";
                else if (af.IdConcepto == 10988) //10988_BANORTE
                    celda = celdaMes + "40";
                else if (af.IdConcepto == 10989) //10989_BANORTE TASA
                    celda = celdaMes + "41";
                else if (af.IdConcepto == 10990) //10990_BANORTE INTERESES PAGADOS
                    celda = celdaMes + "42";
                else if (af.IdConcepto == 10991) //10991_BANSI
                    celda = celdaMes + "43";
                else if (af.IdConcepto == 10992) //10992_BANSI TASA
                    celda = celdaMes + "44";
                else if (af.IdConcepto == 10993) //10993_BANSI INTERESES PAGADOS
                    celda = celdaMes + "45";
                else if (af.IdConcepto == 10994) //10994_BBVA
                    celda = celdaMes + "46";
                else if (af.IdConcepto == 10995) //10995_BBVA TASA
                    celda = celdaMes + "47";
                else if (af.IdConcepto == 10996) //10996_BBVA INTERESES PAGADOS
                    celda = celdaMes + "48";
                else if (af.IdConcepto == 10997) //10997_CLEAR LEASING
                    celda = celdaMes + "49";
                else if (af.IdConcepto == 10998) //10998_CLEAR LEASING TASA
                    celda = celdaMes + "50";
                else if (af.IdConcepto == 10999) //10999_CLEAR LEASING INTERESES PAGADOS
                    celda = celdaMes + "51";
                else if (af.IdConcepto == 11000) //11000_DISI OPERACIONES
                    celda = celdaMes + "52";
                else if (af.IdConcepto == 11001) //11001_DISI OPERACIONES TASA
                    celda = celdaMes + "53";
                else if (af.IdConcepto == 11002) //11002_DISI OPERACIONES INTERESES PAGADOS
                    celda = celdaMes + "54";
                else if (af.IdConcepto == 11003) //11003_EXITUS
                    celda = celdaMes + "55";
                else if (af.IdConcepto == 11004) //11004_EXITUS TASA
                    celda = celdaMes + "56";
                else if (af.IdConcepto == 11005) //11005_EXITUS INTERESES PAGADOS
                    celda = celdaMes + "57";
                else if (af.IdConcepto == 11006) //11006_IMPORTADORA COOLSALES
                    celda = celdaMes + "58";
                else if (af.IdConcepto == 11007) //11007_IMPORTADORA COOLSALES TASA
                    celda = celdaMes + "59";
                else if (af.IdConcepto == 11008) //11008_IMPORTADORA COOLSALES INTERESES PAGADOS
                    celda = celdaMes + "60";
                else if (af.IdConcepto == 11009) //11009_INBURSA
                    celda = celdaMes + "61";
                else if (af.IdConcepto == 11010) //11010_INBURSA TASA
                    celda = celdaMes + "62";
                else if (af.IdConcepto == 11011) //11011_INBURSA INTERESES PAGADOS
                    celda = celdaMes + "63";
                else if (af.IdConcepto == 11012) //11012_NR FINANCE
                    celda = celdaMes + "64";
                else if (af.IdConcepto == 11013) //11013_NR FINANCE TASA
                    celda = celdaMes + "65";
                else if (af.IdConcepto == 11014) //11014_NR FINANCE INTERESES PAGADOS
                    celda = celdaMes + "66";
                else if (af.IdConcepto == 11015) //11015_VOLKSWAGEN
                    celda = celdaMes + "67";
                else if (af.IdConcepto == 11016) //11016_VOLKSWAGEN TASA
                    celda = celdaMes + "68";
                else if (af.IdConcepto == 11017) //11017_VOLKSWAGEN INTERESES PAGADOS
                    celda = celdaMes + "69";
                else if (af.IdConcepto == 11018) //11018_CARTAS DE CRÉDITO BANBAJIO 1
                    celda = celdaMes + "70";
                else if (af.IdConcepto == 11019) //11019_CARTAS DE CRÉDITO BANBAJIO 2
                    celda = celdaMes + "71";
                else if (af.IdConcepto == 11020) //11020_CARTAS DE CRÉDITO BANBAJIO 3
                    celda = celdaMes + "72";
                else if (af.IdConcepto == 11021) //11021_CARTAS DE CRÉDITO SANTANDER
                    celda = celdaMes + "73";
                else if (af.IdConcepto == 11022) //11022_CARTAS DE CRÉDITO SANTANDER TASA
                    celda = celdaMes + "74";
                else if (af.IdConcepto == 11023) //11023_CARTAS DE CRÉDITO SANTANDER INTERESES PAGADOS
                    celda = celdaMes + "75";
                else if (af.IdConcepto == 11024) //11024_BMW FINANCIAL
                    celda = celdaMes + "76";
                else if (af.IdConcepto == 11025) //11025_BMW FINANCIAL TASA
                    celda = celdaMes + "77";
                else if (af.IdConcepto == 11026) //11026_BMW FINANCIAL INTERESES PAGADOS
                    celda = celdaMes + "78";
                else if (af.IdConcepto == 127) //127_CUFIN AL 31 DE DICIEMBRE 2017
                    celda = celdaMes + "90";
                else if (af.IdConcepto == 128) //128_DIVIDENDOS PAGADOS
                    celda = celdaMes + "91";
                else if (af.IdConcepto == 129) //129_SU PRESTADORA
                    celda = celdaMes + "97";
                else if (af.IdConcepto == 130) //130_INMOBILIARIA ARRULLO
                    celda = celdaMes + "98";
                else if (af.IdConcepto == 131) //131_GMC BUICK
                    celda = celdaMes + "99";

                if (af.IdConcepto == 119 || af.IdConcepto == 10416 || af.IdConcepto == 10418 || af.IdConcepto == 124 || af.IdConcepto == 126)
                    banEsPorcentaje = true;

                if (celda.Length >= 2)
                {
                    if (mes == 1)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Ene / 100) : af.Ene;
                    else if (mes == 2)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Feb / 100) : af.Feb;
                    else if (mes == 3)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Mar / 100) : af.Mar;
                    else if (mes == 4)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Abr / 100) : af.Abr;
                    else if (mes == 5)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.May / 100) : af.May;
                    else if (mes == 6)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Jun / 100) : af.Jun;
                    else if (mes == 7)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Jul / 100) : af.Jul;
                    else if (mes == 8)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Ago / 100) : af.Ago;
                    else if (mes == 9)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Sep / 100) : af.Sep;
                    else if (mes == 10)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Oct / 100) : af.Oct;
                    else if (mes == 11)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Nov / 100) : af.Nov;
                    else if (mes == 12)
                        ws.get_Range(celda, celda).Formula = banEsPorcentaje ? (af.Dic / 100) : af.Dic;
                }

                celda = "";
                banEsPorcentaje = false;

                Console.WriteLine("[CONCEPTO]: " + af.IdConcepto + "_" + af.Concepto);
            }
        }

        public void LlenaPestaniaRMANyRMAU(ExcelApp.Worksheet ws)
        {
            string celda = "";

            //celda = "B1";

            //if (celda == "B1") //Razón Social
            //    ws.get_Range(celda, celda).Formula = razonSocial;

            //celda = "B32";

            //if (celda == "B32") //Razón Social
            //    ws.get_Range(celda, celda).Formula = razonSocial;

            //celda = "";

            if (version == "V1")
            {
                List<ResultadoMesAutosNuevos> listRMANV1 = ResultadoMesAutosNuevos.Listar(_db, idAgencia, anio, 1, true);

                foreach (ResultadoMesAutosNuevos rman in listRMANV1)
                {
                    celda = GetMes();

                    if (rman.IdConcepto == 1000 && rman.IdGrupo == 1) //1000_INGRESOS, VENTA TRADICIONAL
                        celda += "4";
                    else if (rman.IdConcepto == 1001 && rman.IdGrupo == 1) //1001_UNIDADES VENDIDAS TRADICIONAL
                        celda += "5";
                    else if (rman.IdConcepto == 1004 && rman.IdGrupo == 1) //1004_UTILIDAD BRUTA TRADICIONAL
                        celda += "7";
                    else if (rman.IdConcepto == 1005 && rman.IdGrupo == 1) //1005_GASTOS DEPARTAMENTALES TRADICIONAL
                        celda += "8";
                    else if (rman.IdConcepto == 1007 && rman.IdGrupo == 1) //1007_UTILIDAD NETA SERVICIOS ADICIONALES TRADICIONAL
                        celda += "10";
                    else if (rman.IdConcepto == 1000 && rman.IdGrupo == 2) //1000_INGRESOS, VENTA DE FLOTILLAS
                        celda += "12";
                    else if (rman.IdConcepto == 1001 && rman.IdGrupo == 2) //1001_UNIDADES VENDIDAS FLOTILLAS
                        celda += "13";
                    else if (rman.IdConcepto == 1004 && rman.IdGrupo == 2) //1004_UTILIDAD BRUTA FLOTILLAS
                        celda += "15";
                    else if (rman.IdConcepto == 1005 && rman.IdGrupo == 2) //1005_GASTOS DEPARTAMENTALES FLOTILLAS
                        celda += "16";
                    else if (rman.IdConcepto == 1007 && rman.IdGrupo == 2) //1007_UTILIDAD NETA SERVICIOS ADICIONALES FLOTILLAS
                        celda += "18";
                    else if (rman.IdConcepto == 1000 && rman.IdGrupo == 3) //1000_INGRESOS, VENTAS POR INTERNET
                        celda += "20";
                    else if (rman.IdConcepto == 1001 && rman.IdGrupo == 3) //1001_UNIDADES VENDIDAS INTERNET
                        celda += "21";
                    else if (rman.IdConcepto == 1004 && rman.IdGrupo == 3) //1004_UTILIDAD BRUTA INTERNET
                        celda += "23";
                    else if (rman.IdConcepto == 1005 && rman.IdGrupo == 3) //1005_GASTOS DEPARTAMENTALES INTERNET
                        celda += "24";
                    else if (rman.IdConcepto == 1007 && rman.IdGrupo == 3) //1007_UTILIDAD NETA SERVICIOS ADICIONALES INTERNET
                        celda += "26";

                    if (celda.Length >= 2)
                    {
                        if (mes == 1)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Enero);
                        else if (mes == 2)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Febrero);
                        else if (mes == 3)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Marzo);
                        else if (mes == 4)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Abril);
                        else if (mes == 5)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Mayo);
                        else if (mes == 6)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Junio);
                        else if (mes == 7)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Julio);
                        else if (mes == 8)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Agosto);
                        else if (mes == 9)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Septiembre);
                        else if (mes == 10)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Octubre);
                        else if (mes == 11)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Noviembre);
                        else if (mes == 12)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Diciembre);
                    }

                    celda = "";

                    Console.WriteLine("[CONCEPTO] RMANV1: " + rman.IdConcepto + "_" + rman.IdGrupo);
                }

                List<ResultadoMesAutosNuevos> listRMAUV1 = ResultadoMesAutosNuevos.Listar(_db, idAgencia, anio, 1, false);

                foreach (ResultadoMesAutosNuevos rmau in listRMAUV1)
                {
                    celda = GetMes();

                    if (rmau.IdConcepto == 1000 && rmau.IdGrupo == 1) //1000_INGRESOS, VENTA TRADICIONAL
                        celda += "35";
                    else if (rmau.IdConcepto == 1013 && rmau.IdGrupo == 1) //1013_UNIDADES VENDIDAS TRADICIONAL
                        celda += "36";
                    else if (rmau.IdConcepto == 1016 && rmau.IdGrupo == 1) //1016_UTILIDAD BRUTA TRADICIONAL
                        celda += "38";
                    else if (rmau.IdConcepto == 1017 && rmau.IdGrupo == 1) //1017_GASTOS DEPARTAMENTALES TRADICIONAL
                        celda += "39";
                    else if (rmau.IdConcepto == 1019 && rmau.IdGrupo == 1) //1019_UTILIDAD NETA SERVICIOS ADICIONALES TRADICIONAL
                        celda += "41";
                    else if (rmau.IdConcepto == 1000 && rmau.IdGrupo == 2) //1000_INGRESOS, VENTA DE FLOTILLAS
                        celda += "43";
                    else if (rmau.IdConcepto == 1013 && rmau.IdGrupo == 2) //1013_UNIDADES VENDIDAS FLOTILLAS
                        celda += "44";
                    else if (rmau.IdConcepto == 1016 && rmau.IdGrupo == 2) //1016_UTILIDAD BRUTA FLOTILLAS
                        celda += "46";
                    else if (rmau.IdConcepto == 1017 && rmau.IdGrupo == 2) //1017_GASTOS DEPARTAMENTALES FLOTILLAS
                        celda += "47";
                    else if (rmau.IdConcepto == 1019 && rmau.IdGrupo == 2) //1019_UTILIDAD NETA SERVICIOS ADICIONALES FLOTILLAS
                        celda += "49";
                    else if (rmau.IdConcepto == 1000 && rmau.IdGrupo == 3) //1000_INGRESOS, VENTAS POR INTERNET
                        celda += "51";
                    else if (rmau.IdConcepto == 1013 && rmau.IdGrupo == 3) //1013_UNIDADES VENDIDAS INTERNET
                        celda += "52";
                    else if (rmau.IdConcepto == 1016 && rmau.IdGrupo == 3) //1016_UTILIDAD BRUTA INTERNET
                        celda += "54";
                    else if (rmau.IdConcepto == 1017 && rmau.IdGrupo == 3) //1017_GASTOS DEPARTAMENTALES INTERNET
                        celda += "55";
                    else if (rmau.IdConcepto == 1019 && rmau.IdGrupo == 3) //1019_UTILIDAD NETA SERVICIOS ADICIONALES INTERNET
                        celda += "57";

                    if (celda.Length >= 2)
                    {
                        if (mes == 1)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Enero);
                        else if (mes == 2)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Febrero);
                        else if (mes == 3)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Marzo);
                        else if (mes == 4)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Abril);
                        else if (mes == 5)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Mayo);
                        else if (mes == 6)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Junio);
                        else if (mes == 7)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Julio);
                        else if (mes == 8)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Agosto);
                        else if (mes == 9)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Septiembre);
                        else if (mes == 10)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Octubre);
                        else if (mes == 11)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Noviembre);
                        else if (mes == 12)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Diciembre);
                    }

                    celda = "";

                    Console.WriteLine("[CONCEPTO] RMAUV1: " + rmau.IdConcepto + "_" + rmau.IdGrupo);
                }
            }
            else if (version == "V2")
            {
                List<ResultadoMesAutosNuevos> listRMANV2 = ResultadoMesAutosNuevos.ListarExtralibros(_db, idAgencia, anio, 1, true);

                foreach (ResultadoMesAutosNuevos rman in listRMANV2)
                {
                    celda = GetMes();

                    if (rman.IdConcepto == 1000 && rman.IdGrupo == 1) //1000_INGRESOS, VENTA TRADICIONAL
                        celda += "4";
                    else if (rman.IdConcepto == 1001 && rman.IdGrupo == 1) //1001_UNIDADES VENDIDAS TRADICIONAL
                        celda += "5";
                    else if (rman.IdConcepto == 1004 && rman.IdGrupo == 1) //1004_UTILIDAD BRUTA TRADICIONAL
                        celda += "7";
                    else if (rman.IdConcepto == 1005 && rman.IdGrupo == 1) //1005_GASTOS DEPARTAMENTALES TRADICIONAL
                        celda += "8";
                    else if (rman.IdConcepto == 1007 && rman.IdGrupo == 1) //1007_UTILIDAD NETA SERVICIOS ADICIONALES TRADICIONAL
                        celda += "10";
                    else if (rman.IdConcepto == 1000 && rman.IdGrupo == 2) //1000_INGRESOS, VENTA DE FLOTILLAS
                        celda += "12";
                    else if (rman.IdConcepto == 1001 && rman.IdGrupo == 2) //1001_UNIDADES VENDIDAS FLOTILLAS
                        celda += "13";
                    else if (rman.IdConcepto == 1004 && rman.IdGrupo == 2) //1004_UTILIDAD BRUTA FLOTILLAS
                        celda += "15";
                    else if (rman.IdConcepto == 1005 && rman.IdGrupo == 2) //1005_GASTOS DEPARTAMENTALES FLOTILLAS
                        celda += "16";
                    else if (rman.IdConcepto == 1007 && rman.IdGrupo == 2) //1007_UTILIDAD NETA SERVICIOS ADICIONALES FLOTILLAS
                        celda += "18";
                    else if (rman.IdConcepto == 1000 && rman.IdGrupo == 3) //1000_INGRESOS, VENTAS POR INTERNET
                        celda += "20";
                    else if (rman.IdConcepto == 1001 && rman.IdGrupo == 3) //1001_UNIDADES VENDIDAS INTERNET
                        celda += "21";
                    else if (rman.IdConcepto == 1004 && rman.IdGrupo == 3) //1004_UTILIDAD BRUTA INTERNET
                        celda += "23";
                    else if (rman.IdConcepto == 1005 && rman.IdGrupo == 3) //1005_GASTOS DEPARTAMENTALES INTERNET
                        celda += "24";
                    else if (rman.IdConcepto == 1007 && rman.IdGrupo == 3) //1007_UTILIDAD NETA SERVICIOS ADICIONALES INTERNET
                        celda += "26";

                    if (celda.Length >= 2)
                    {
                        if (mes == 1)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Enero);
                        else if (mes == 2)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Febrero);
                        else if (mes == 3)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Marzo);
                        else if (mes == 4)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Abril);
                        else if (mes == 5)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Mayo);
                        else if (mes == 6)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Junio);
                        else if (mes == 7)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Julio);
                        else if (mes == 8)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Agosto);
                        else if (mes == 9)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Septiembre);
                        else if (mes == 10)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Octubre);
                        else if (mes == 11)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Noviembre);
                        else if (mes == 12)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rman.Diciembre);
                    }

                    celda = "";

                    Console.WriteLine("[CONCEPTO] RMANV2: " + rman.IdConcepto + "_" + rman.IdGrupo);
                }

                List<ResultadoMesAutosNuevos> listRMAUV2 = ResultadoMesAutosNuevos.ListarExtralibros(_db, idAgencia, anio, 1, false);

                foreach (ResultadoMesAutosNuevos rmau in listRMAUV2)
                {
                    celda = GetMes();

                    if (rmau.IdConcepto == 1000 && rmau.IdGrupo == 1) //1000_INGRESOS, VENTA TRADICIONAL
                        celda += "35";
                    else if (rmau.IdConcepto == 1013 && rmau.IdGrupo == 1) //1013_UNIDADES VENDIDAS TRADICIONAL
                        celda += "36";
                    else if (rmau.IdConcepto == 1016 && rmau.IdGrupo == 1) //1016_UTILIDAD BRUTA TRADICIONAL
                        celda += "38";
                    else if (rmau.IdConcepto == 1017 && rmau.IdGrupo == 1) //1017_GASTOS DEPARTAMENTALES TRADICIONAL
                        celda += "39";
                    else if (rmau.IdConcepto == 1019 && rmau.IdGrupo == 1) //1019_UTILIDAD NETA SERVICIOS ADICIONALES TRADICIONAL
                        celda += "41";
                    else if (rmau.IdConcepto == 1000 && rmau.IdGrupo == 2) //1000_INGRESOS, VENTA DE FLOTILLAS
                        celda += "43";
                    else if (rmau.IdConcepto == 1013 && rmau.IdGrupo == 2) //1013_UNIDADES VENDIDAS FLOTILLAS
                        celda += "44";
                    else if (rmau.IdConcepto == 1016 && rmau.IdGrupo == 2) //1016_UTILIDAD BRUTA FLOTILLAS
                        celda += "46";
                    else if (rmau.IdConcepto == 1017 && rmau.IdGrupo == 2) //1017_GASTOS DEPARTAMENTALES FLOTILLAS
                        celda += "47";
                    else if (rmau.IdConcepto == 1019 && rmau.IdGrupo == 2) //1019_UTILIDAD NETA SERVICIOS ADICIONALES FLOTILLAS
                        celda += "49";
                    else if (rmau.IdConcepto == 1000 && rmau.IdGrupo == 3) //1000_INGRESOS, VENTAS POR INTERNET
                        celda += "51";
                    else if (rmau.IdConcepto == 1013 && rmau.IdGrupo == 3) //1013_UNIDADES VENDIDAS INTERNET
                        celda += "52";
                    else if (rmau.IdConcepto == 1016 && rmau.IdGrupo == 3) //1016_UTILIDAD BRUTA INTERNET
                        celda += "54";
                    else if (rmau.IdConcepto == 1017 && rmau.IdGrupo == 3) //1017_GASTOS DEPARTAMENTALES INTERNET
                        celda += "55";
                    else if (rmau.IdConcepto == 1019 && rmau.IdGrupo == 3) //1019_UTILIDAD NETA SERVICIOS ADICIONALES INTERNET
                        celda += "57";

                    if (celda.Length >= 2)
                    {
                        if (mes == 1)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Enero);
                        else if (mes == 2)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Febrero);
                        else if (mes == 3)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Marzo);
                        else if (mes == 4)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Abril);
                        else if (mes == 5)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Mayo);
                        else if (mes == 6)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Junio);
                        else if (mes == 7)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Julio);
                        else if (mes == 8)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Agosto);
                        else if (mes == 9)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Septiembre);
                        else if (mes == 10)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Octubre);
                        else if (mes == 11)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Noviembre);
                        else if (mes == 12)
                            ws.get_Range(celda, celda).Formula = Convert.ToInt32(rmau.Diciembre);
                    }

                    celda = "";

                    Console.WriteLine("[CONCEPTO] RMAUV2: " + rmau.IdConcepto + "_" + rmau.IdGrupo);
                }
            }
        }

        public void LlenaPestaniaRM(ExcelApp.Worksheet ws)
        {
            string celda = "";

            //celda = "B1";

            //if (celda == "B1") //Razón Social
            //    ws.get_Range(celda, celda).Formula = razonSocial;
            
            //celda = "";

            List<ConceptosContables> conceptosV1 = ConceptosContables.ListarRMV1(_db);
            List<ConceptosContables> conceptosV2 = ConceptosContables.ListarRMV2(_db);

            System.Data.DataTable dtRMWebV1 = proc.GetRMWebV1();
            System.Data.DataTable dtRMWebV1Acum = proc.GetRMWebV1Acumulado();
            System.Data.DataTable dtRMWebV2 = proc.GetRMWebV2();
            System.Data.DataTable dtRMWebV2Acum = proc.GetRMWebV2Acumulado();

            int r = 1;
            int vWeb = 0;
            int intAux = 0;
            DataRow[] drRMWebV1 = null;
            DataRow[] drRMWebV1Acum = null;
            DataRow[] drRMWebV2 = null;
            DataRow[] drRMWebV2Acum = null;

            foreach (ConceptosContables concepto in conceptosV2)
            {
                Reporte rep = new Reporte();

                r++;

                rep.ID_CONCEPTO = concepto.Id;
                rep.CONCEPTO = concepto.NombreConcepto;
                                
                if (idAgencia == 27)
                {
                    drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 286)
                {
                    drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 36)
                {
                    drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 12)
                {
                    drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 35)
                {
                    drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 588)
                {
                    drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 32)
                {
                    drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 212)
                {
                    drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                }
                else
                {
                    drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                }
                                
                if (idAgencia == 27)
                {
                    drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 286)
                {
                    drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 36)
                {
                    drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 12)
                {
                    drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 35)
                {
                    drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 588)
                {
                    drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 32)
                {
                    drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                }
                else if (idAgencia == 212)
                {
                    drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                }
                else
                {
                    drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                    drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                }
                
                if (drRMWebV1.Length != 0)
                {
                    vWeb = 0;

                    foreach (DataRow dr in drRMWebV1)
                    {
                        if ((concepto.Id == 1001) || (concepto.Id == 1003) || (concepto.Id == 1013) || (concepto.Id == 1015) || (concepto.Id == 1031)
                             || (concepto.Id == 1032) || (concepto.Id == 1060) || (concepto.Id == 1064) || (concepto.Id == 1065))
                            vWeb += Convert.ToInt32(dr["VALOR"]);
                        else if (concepto.Id == 1033)
                        {
                            FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, idAgencia, FINA_EParametros.HorasFacturadasServVal);

                            if (resp1 != null && int.TryParse(resp1.Valor, out intAux))
                            {
                                intAux = intAux == 0 ? 1 : intAux;

                                vWeb += Convert.ToInt32(dr["VALOR"]) / intAux;
                            }
                        }
                        else
                            vWeb += Convert.ToInt32(dr["VALOR"]) / 1000;
                    }
                    rep.WEB_V1 = vWeb;

                    vWeb = 0;

                    foreach (DataRow dr in drRMWebV1Acum)
                    {
                        if ((concepto.Id == 1001) || (concepto.Id == 1003) || (concepto.Id == 1013) || (concepto.Id == 1015) || (concepto.Id == 1031)
                             || (concepto.Id == 1032) || (concepto.Id == 1060) || (concepto.Id == 1064) || (concepto.Id == 1065))
                            vWeb += Convert.ToInt32(dr["VALOR"]);
                        else if (concepto.Id == 1033)
                        {
                            FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, idAgencia, FINA_EParametros.HorasFacturadasServVal);

                            if (resp1 != null && int.TryParse(resp1.Valor, out intAux))
                            {
                                intAux = intAux == 0 ? 1 : intAux;

                                vWeb += Convert.ToInt32(dr["VALOR"]) / intAux;
                            }
                        }
                        else
                            vWeb += Convert.ToInt32(dr["VALOR"]) / 1000;
                    }
                    rep.WEB_V1_ACUM = vWeb;
                }

                rep.DIFF_V1 = rep.EXCEL_V1 - rep.WEB_V1;
                rep.DIFF_V1_ACUM = rep.EXCEL_V1_ACUM - rep.WEB_V1_ACUM;
                                
                if (drRMWebV2.Length != 0)
                {
                    vWeb = 0;

                    foreach (DataRow dr in drRMWebV2)
                    {
                        if ((concepto.Id == 1001) || (concepto.Id == 1003) || (concepto.Id == 1013) || (concepto.Id == 1015) || (concepto.Id == 1031)
                             || (concepto.Id == 1032) || (concepto.Id == 1060) || (concepto.Id == 1064) || (concepto.Id == 1065))
                            vWeb += Convert.ToInt32(dr["VALOR"]);
                        else if (concepto.Id == 1033)
                        {
                            FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, idAgencia, FINA_EParametros.HorasFacturadasServVal);

                            if (resp1 != null && int.TryParse(resp1.Valor, out intAux))
                            {
                                intAux = intAux == 0 ? 1 : intAux;

                                vWeb += Convert.ToInt32(dr["VALOR"]) / intAux;
                            }
                        }
                        else
                            vWeb += Convert.ToInt32(dr["VALOR"]) / 1000;
                    }
                    rep.WEB_V2 = vWeb;

                    vWeb = 0;

                    foreach (DataRow dr in drRMWebV2Acum)
                    {
                        if ((concepto.Id == 1001) || (concepto.Id == 1003) || (concepto.Id == 1013) || (concepto.Id == 1015) || (concepto.Id == 1031)
                             || (concepto.Id == 1032) || (concepto.Id == 1060) || (concepto.Id == 1064) || (concepto.Id == 1065))
                            vWeb += Convert.ToInt32(dr["VALOR"]);
                        else if (concepto.Id == 1033)
                        {
                            FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, idAgencia, FINA_EParametros.HorasFacturadasServVal);

                            if (resp1 != null && int.TryParse(resp1.Valor, out intAux))
                            {
                                intAux = intAux == 0 ? 1 : intAux;

                                vWeb += Convert.ToInt32(dr["VALOR"]) / intAux;
                            }
                        }
                        else
                            vWeb += Convert.ToInt32(dr["VALOR"]) / 1000;
                    }
                    rep.WEB_V2_ACUM = vWeb;
                }

                rep.DIFF_V2 = rep.EXCEL_V2 - rep.WEB_V2;
                rep.DIFF_V2_ACUM = rep.EXCEL_V2_ACUM - rep.WEB_V2_ACUM;

                celda = GetMes();

                if (version == "V1")
                {
                    if (concepto.Id == 1009) //1009_INGRESOS GERENCIA DE NEGOCIOS
                        celda += "13";
                    else if (concepto.Id == 1010) //1010_UTILIDAD BRUTA DEPARTAMENTAL GERENCIA DE NEGOCIOS
                        celda += "14";
                    else if (concepto.Id == 1011) //1011_GASTOS DEPARTAMENTALES GERENCIA DE NEGOCIOS
                        celda += "15";
                    else if (concepto.Id == 1021) //1021_INGRESOS POR VENTA DE SERVICIO
                        celda += "25";
                    else if (concepto.Id == 1022) //1022_UTILIDAD BRUTA DEPARTAMENTAL SERVICIO
                        celda += "26";
                    else if (concepto.Id == 1023) //1023_GASTOS DEPARTAMENTALES SERVICIO
                        celda += "27";
                    else if (concepto.Id == 1025) //1025_INGRESOS POR HOJALATERIA Y PINTURA
                        celda += "29";
                    else if (concepto.Id == 1026) //1026_UTILIDAD BRUTA DEPARTAMENTAL HOJALATERIA Y PINTURA
                        celda += "30";
                    else if (concepto.Id == 1027) //1027_GASTOS DEPARTAMENTALES HOJALATERIA Y PINTURA
                        celda += "31";
                    else if (concepto.Id == 1029) //1029_UTILIDAD NETA SERVICIOS ADICIONALES HOJALATERIA Y PINTURA
                        celda += "33";
                    else if (concepto.Id == 1031) //1031_No. FACTURAS EMITIDAS SERVICIO
                        celda += "35";
                    else if (concepto.Id == 1032) //1032_No. FACTURAS EMITIDAS HYP
                        celda += "36";
                    else if (concepto.Id == 1033) //1033_No. HORAS FACTURADAS SERVICIO
                        celda += "37";
                    else if (concepto.Id == 1035) //1035_VENTA DE REFACCIONES SERVICIO
                        celda += "39";
                    else if (concepto.Id == 1036) //1036_VENTA DE REFACCIONES H&P
                        celda += "40";
                    else if (concepto.Id == 1037) //1037_VENTA MAYOREO
                        celda += "41";
                    else if (concepto.Id == 1038) //1038_VENTA MOSTRADOR
                        celda += "42";
                    else if (concepto.Id == 1039) //1039_UTILIDAD BRUTA DEPARTAMENTAL REFACCIONES
                        celda += "43";
                    else if (concepto.Id == 1040) //1040_GASTOS DEPARTAMENTALES REFACCIONES
                        celda += "44";
                    else if (concepto.Id == 1042) //1042_UTILIDAD NETA SERVICIOS ADICIONALES REFACCIONES
                        celda += "46";
                    else if (concepto.Id == 1044) //1044_GASTOS ADMINISTRATIVOS
                        celda += "48";
                    else if (concepto.Id == 1045) //1045_OTROS INGRESOS PLANTA
                        celda += "49";
                    else if (concepto.Id == 1052) //1052_UTILIDAD NETA DEL FIDEICOMISO
                        celda += "56";
                    else if (concepto.Id == 1053) //1053_UTILIDAD NETA TAXIS BAM
                        celda += "57";
                    else if (concepto.Id == 1058) //1058_P.T.U.
                        celda += "62";
                }
                else
                {
                    if (concepto.Id == 1088) //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS
                        celda += "13";
                    else if (concepto.Id == 1009) //1009_INGRESOS GERENCIA DE NEGOCIOS
                        celda += "15";
                    else if (concepto.Id == 1010) //1010_UTILIDAD BRUTA DEPARTAMENTAL GERENCIA DE NEGOCIOS
                        celda += "16";
                    else if (concepto.Id == 1011) //1011_GASTOS DEPARTAMENTALES GERENCIA DE NEGOCIOS
                        celda += "17";
                    else if (concepto.Id == 1089) //1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS
                        celda += "19";
                    else if (concepto.Id == 1090) //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS
                        celda += "29";
                    else if (concepto.Id == 1021) //1021_INGRESOS POR VENTA DE SERVICIO
                        celda += "31";
                    else if (concepto.Id == 1022) //1022_UTILIDAD BRUTA DEPARTAMENTAL SERVICIO
                        celda += "32";
                    else if (concepto.Id == 1023) //1023_GASTOS DEPARTAMENTALES SERVICIO
                        celda += "33";
                    else if (concepto.Id == 1091) //1091_GASTOS DE ADMINISTRACION SERVICIO
                        celda += "35";
                    else if (concepto.Id == 1025) //1025_INGRESOS POR HOJALATERIA Y PINTURA
                        celda += "37";
                    else if (concepto.Id == 1026) //1026_UTILIDAD BRUTA DEPARTAMENTAL HOJALATERIA Y PINTURA
                        celda += "38";
                    else if (concepto.Id == 1027) //1027_GASTOS DEPARTAMENTALES HOJALATERIA Y PINTURA
                        celda += "39";
                    else if (concepto.Id == 1092) //1092_GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA
                        celda += "41";
                    else if (concepto.Id == 1029) //1029_UTILIDAD NETA SERVICIOS ADICIONALES HOJALATERIA Y PINTURA
                        celda += "43";
                    else if (concepto.Id == 1031) //1031_No. FACTURAS EMITIDAS SERVICIO
                        celda += "45";
                    else if (concepto.Id == 1032) //1032_No. FACTURAS EMITIDAS HYP
                        celda += "46";
                    else if (concepto.Id == 1033) //1033_No. HORAS FACTURADAS SERVICIO
                        celda += "47";
                    else if (concepto.Id == 1035) //1035_VENTA DE REFACCIONES SERVICIO
                        celda += "49";
                    else if (concepto.Id == 1036) //1036_VENTA DE REFACCIONES H&P
                        celda += "50";
                    else if (concepto.Id == 1037) //1037_VENTA MAYOREO
                        celda += "51";
                    else if (concepto.Id == 1038) //1038_VENTA MOSTRADOR
                        celda += "52";
                    else if (concepto.Id == 1039) //1039_UTILIDAD BRUTA DEPARTAMENTAL REFACCIONES
                        celda += "53";
                    else if (concepto.Id == 1040) //1040_GASTOS DEPARTAMENTALES REFACCIONES
                        celda += "54";
                    else if (concepto.Id == 1042) //1042_UTILIDAD NETA SERVICIOS ADICIONALES REFACCIONES
                        celda += "55";
                    else if (concepto.Id == 1093) //1093_GASTOS DE ADMINISTRACION REFACCIONES
                        celda += "58";
                    else if (concepto.Id == 1044) //1044_GASTOS ADMINISTRATIVOS
                        celda += "77";
                    else if (concepto.Id == 1045) //1045_OTROS INGRESOS PLANTA
                        celda += "60";
                    else if (concepto.Id == 1052) //1052_UTILIDAD NETA DEL FIDEICOMISO
                        celda += "67";
                    else if (concepto.Id == 1053) //1053_UTILIDAD NETA TAXIS BAM
                        celda += "68";
                    else if (concepto.Id == 1058) //1058_P.T.U.
                        celda += "73";
                }               

                if (celda.Length >= 2)
                {
                    if (version == "V1")
                        ws.get_Range(celda, celda).Formula = rep.WEB_V1;
                    else
                        ws.get_Range(celda, celda).Formula = rep.WEB_V2;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + concepto.Id + "_" + concepto.NombreConcepto);
            }
        }

        public void LlenaPestaniaOPL(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            //celda = "A1";

            //if (celda == "A1") //Razón Social
            //    ws.get_Range(celda, celda).Formula = razonSocial;

            //celda = "";

            int finanNetoAN = 0;
            int finanNetoAS = 0;
            int finanNetoRE = 0;
            int finanNeto = 0;
            int otroGastosProd = 0;
            int gastoCorpo = 0;
            int isrCorriente = 0;
            int isrDiferido = 0;

            List<CapturaOPL> lstOPL = CapturaOPL.ListarAnioMes(_db, idAgencia, anio.ToString(), mes);

            foreach (CapturaOPL opl in lstOPL)
            {
                if (opl.IdGrupo == 1 && opl.IdMaestro == 887)
                    finanNetoAN = finanNetoAN + Convert.ToInt32(opl.Valor);
                else if (opl.IdGrupo == 1 && opl.IdMaestro == 888)
                    finanNetoAS = finanNetoAS + Convert.ToInt32(opl.Valor);
                else if (opl.IdGrupo == 1 && opl.IdMaestro == 889)
                    finanNetoRE = finanNetoRE + Convert.ToInt32(opl.Valor);
                else if (opl.IdGrupo == 1 && (opl.IdMaestro != 887 || opl.IdMaestro != 888 || opl.IdMaestro != 889))
                    finanNeto = finanNeto + Convert.ToInt32(opl.Valor);
                else if (opl.IdGrupo == 2)
                    otroGastosProd = otroGastosProd + Convert.ToInt32(opl.Valor);
                else if (opl.IdGrupo == 3)
                    gastoCorpo = gastoCorpo + Convert.ToInt32(opl.Valor);
                else if (opl.IdGrupo == 4)
                    isrCorriente = isrCorriente + Convert.ToInt32(opl.Valor);
                else if (opl.IdGrupo == 5)
                    isrDiferido = isrDiferido + Convert.ToInt32(opl.Valor);
            }

            PeriodoContable periodo = PeriodoContable.BuscarPorMesAnio(_db, mes, anio);

            if (finanNetoAN == 0)
            {
                SaldoPorPeriodoPorCuentaBalanza saldo621200010001 = SaldoPorPeriodoPorCuentaBalanza.Buscar(_db, idAgencia, periodo.Id, 1, 23946); //23946_621200010001_GASTOS FINANCIEROS || INTERESES PLAN PISO PLANTA || INTERESES PLAN PISO PLANTA ||
                SaldoPorPeriodoPorCuentaBalanza saldo621200010002 = SaldoPorPeriodoPorCuentaBalanza.Buscar(_db, idAgencia, periodo.Id, 1, 23947); //23947_621200010002_GASTOS FINANCIEROS || INTERESES PLAN PISO PLANTA || BONIF.PLAN PISO PLANTA ||
                SaldoPorPeriodoPorCuentaBalanza saldo621200010003 = SaldoPorPeriodoPorCuentaBalanza.Buscar(_db, idAgencia, periodo.Id, 1, 304438); //304438_621200010003_GASTOS FINANCIEROS || INTERESES PLAN PISO PLANTA || AUTOS NUEVOS ||

                if (saldo621200010001 != null)
                    finanNetoAN += Convert.ToInt32(saldo621200010001.TotalDeCargos - saldo621200010001.TotalDeAbonos) / 1000;

                if (saldo621200010002 != null)
                    finanNetoAN += Convert.ToInt32(saldo621200010002.TotalDeCargos - saldo621200010002.TotalDeAbonos) / 1000;

                if (saldo621200010003 != null)
                    finanNetoAN += Convert.ToInt32(saldo621200010003.TotalDeCargos - saldo621200010003.TotalDeAbonos) / 1000;

                if (finanNetoAN != 0)
                {
                    finanNetoAN = finanNetoAN * -1;

                    CapturaOPL opl = new CapturaOPL();
                    opl.IdAgencia = idAgencia;
                    opl.IdGrupo = 1;
                    opl.IdConcepto = 10297;
                    opl.IdMaestro = 887;
                    opl.IdMes = mes;
                    opl.Anio = anio.ToString();
                    opl.Valor = finanNetoAN;

                    _db.Insert<CapturaOPL>(3665, 0, opl);
                }
            }

            if (finanNetoAS == 0)
            {
                SaldoPorPeriodoPorCuentaBalanza saldo621200010006 = SaldoPorPeriodoPorCuentaBalanza.Buscar(_db, idAgencia, periodo.Id, 1, 304441); //304441_621200010006_GASTOS FINANCIEROS || INTERESES PLAN PISO PLANTA || AUTOS SEMINUEVOS ||

                if (saldo621200010006 != null)
                {
                    finanNetoAS = Convert.ToInt32(saldo621200010006.TotalDeCargos - saldo621200010006.TotalDeAbonos) / 1000;

                    if (finanNetoAS != 0)
                    {
                        finanNetoAS = finanNetoAS * -1;

                        CapturaOPL opl = new CapturaOPL();
                        opl.IdAgencia = idAgencia;
                        opl.IdGrupo = 1;
                        opl.IdConcepto = 10419;
                        opl.IdMaestro = 888;
                        opl.IdMes = mes;
                        opl.Anio = anio.ToString();
                        opl.Valor = finanNetoAS;

                        _db.Insert<CapturaOPL>(3665, 0, opl);
                    }
                }
            }

            if (finanNetoRE == 0)
            {
                SaldoPorPeriodoPorCuentaBalanza saldo621200010009 = SaldoPorPeriodoPorCuentaBalanza.Buscar(_db, idAgencia, periodo.Id, 1, 304444); //304444_621200010009_GASTOS FINANCIEROS || INTERESES PLAN PISO PLANTA || REFACCIONES ||

                if (saldo621200010009 != null)
                {
                    finanNetoRE = Convert.ToInt32(saldo621200010009.TotalDeCargos - saldo621200010009.TotalDeAbonos) / 1000;

                    if (finanNetoRE != 0)
                    {
                        finanNetoRE = finanNetoRE * -1;

                        CapturaOPL opl = new CapturaOPL();
                        opl.IdAgencia = idAgencia;
                        opl.IdGrupo = 1;
                        opl.IdConcepto = 10420;
                        opl.IdMaestro = 889;
                        opl.IdMes = mes;
                        opl.Anio = anio.ToString();
                        opl.Valor = finanNetoRE;

                        _db.Insert<CapturaOPL>(3665, 0, opl);
                    }
                }
            }

            if (finanNeto == 0)
            {
                SaldoPorPeriodoPorCuentaBalanza saldo62120003 = SaldoPorPeriodoPorCuentaBalanza.Buscar(_db, idAgencia, periodo.Id, 1, 23948); //23948_62120003_GASTOS FINANCIEROS || INTERESES PAGADOS A BANCOS ||
                SaldoPorPeriodoPorCuentaBalanza saldo62120007 = SaldoPorPeriodoPorCuentaBalanza.Buscar(_db, idAgencia, periodo.Id, 1, 23951); //23951_62120007_GASTOS FINANCIEROS || INTERES PAGADOS ORG.HERPA ||

                if (saldo62120003 != null)
                    finanNeto += Convert.ToInt32(saldo62120003.TotalDeCargos - saldo62120003.TotalDeAbonos) / 1000;

                if (saldo62120007 != null)
                    finanNeto += Convert.ToInt32(saldo62120007.TotalDeCargos - saldo62120007.TotalDeAbonos) / 1000;

                if (finanNeto != 0)
                {
                    finanNeto = finanNeto * -1;

                    CapturaOPL opl = new CapturaOPL();
                    opl.IdAgencia = idAgencia;
                    opl.IdGrupo = 1;
                    opl.IdConcepto = 10421;
                    opl.IdMaestro = 890;
                    opl.IdMes = mes;
                    opl.Anio = anio.ToString();
                    opl.Valor = finanNeto;

                    _db.Insert<CapturaOPL>(3665, 0, opl);
                }
            }

            if (isrCorriente == 0)
            {
                ReporteRM rm = ReporteRM.ListarExtralibros(_db, idAgencia, anio).Find(x => x.IdConcepto == 1056);

                CapturaOPL opl = new CapturaOPL();
                opl.IdAgencia = idAgencia;
                opl.IdGrupo = 4;
                opl.IdConcepto = 10231;
                opl.IdMaestro = 1078;
                opl.IdMes = mes;
                opl.Anio = anio.ToString();
                if (mes == 1)
                    opl.Valor = rm.Enevta;
                else if (mes == 2)
                    opl.Valor = rm.Febvta;
                else if (mes == 3)
                    opl.Valor = rm.Marvta;
                else if (mes == 4)
                    opl.Valor = rm.Abrvta;
                else if (mes == 5)
                    opl.Valor = rm.Mayvta;
                else if (mes == 6)
                    opl.Valor = rm.Junvta;
                else if (mes == 7)
                    opl.Valor = rm.Julvta;
                else if (mes == 8)
                    opl.Valor = rm.Agovta;
                else if (mes == 9)
                    opl.Valor = rm.Sepvta;
                else if (mes == 10)
                    opl.Valor = rm.Octvta;
                else if (mes == 11)
                    opl.Valor = rm.Novvta;
                else if (mes == 12)
                    opl.Valor = rm.Dicvta;

                if (opl.Valor != 0)
                {
                    isrCorriente = Convert.ToInt32(opl.Valor);

                    _db.Insert<CapturaOPL>(3665, 0, opl);
                }
            }

            if (isrDiferido == 0)
            {
                ReporteRM rm = ReporteRM.ListarExtralibros(_db, idAgencia, anio).Find(x => x.IdConcepto == 1057);

                CapturaOPL opl = new CapturaOPL();
                opl.IdAgencia = idAgencia;
                opl.IdGrupo = 5;
                opl.IdConcepto = 10698;
                opl.IdMaestro = 1418;
                opl.IdMes = mes;
                opl.Anio = anio.ToString();
                if (mes == 1)
                    opl.Valor = rm.Enevta;
                else if (mes == 2)
                    opl.Valor = rm.Febvta;
                else if (mes == 3)
                    opl.Valor = rm.Marvta;
                else if (mes == 4)
                    opl.Valor = rm.Abrvta;
                else if (mes == 5)
                    opl.Valor = rm.Mayvta;
                else if (mes == 6)
                    opl.Valor = rm.Junvta;
                else if (mes == 7)
                    opl.Valor = rm.Julvta;
                else if (mes == 8)
                    opl.Valor = rm.Agovta;
                else if (mes == 9)
                    opl.Valor = rm.Sepvta;
                else if (mes == 10)
                    opl.Valor = rm.Octvta;
                else if (mes == 11)
                    opl.Valor = rm.Novvta;
                else if (mes == 12)
                    opl.Valor = rm.Dicvta;

                if (opl.Valor != 0)
                {
                    isrDiferido = Convert.ToInt32(opl.Valor);

                    _db.Insert<CapturaOPL>(3665, 0, opl);
                }
            }

            celdaMes = GetMes();
            
            celda = celdaMes + "5"; //1_FINANCIAMIENTO NETO AUTOS NUEVOS
            ws.get_Range(celda, celda).Formula = finanNetoAN;

            Console.WriteLine("[CONCEPTO]: 1_FINANCIAMIENTO NETO AUTOS NUEVOS = " + finanNetoAN);

            celda = celdaMes + "6"; //2_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
            ws.get_Range(celda, celda).Formula = finanNetoAS;

            Console.WriteLine("[CONCEPTO]: 2_FINANCIAMIENTO NETO AUTOS SEMINUEVOS = " + finanNetoAS);

            celda = celdaMes + "7"; //3_FINANCIAMIENTO NETO REFACCIONES
            ws.get_Range(celda, celda).Formula = finanNetoRE;

            Console.WriteLine("[CONCEPTO]: 3_FINANCIAMIENTO NETO REFACCIONES = " + finanNetoRE);

            celda = celdaMes + "8"; //4_FINANCIAMIENTO NETO (PRESTAMOS BANCARIOS PARA TERCEROS)
            ws.get_Range(celda, celda).Formula = finanNeto;

            Console.WriteLine("[CONCEPTO]: 4_FINANCIAMIENTO NETO (PRESTAMOS BANCARIOS PARA TERCEROS) = " + finanNeto);

            celda = celdaMes + "23"; //OTROS (GASTOS) / PRODUCTOS
            ws.get_Range(celda, celda).Formula = otroGastosProd;

            Console.WriteLine("[CONCEPTO]: OTROS (GASTOS) / PRODUCTOS = " + otroGastosProd);

            if (version == "V1")
            {
                celda = celdaMes + "52"; //GASTOS CORPORATIVOS
                ws.get_Range(celda, celda).Formula = gastoCorpo;

                Console.WriteLine("[CONCEPTO]: GASTOS CORPORATIVOS = " + gastoCorpo);

                celda = celdaMes + "70"; //I.S.R. CORRIENTE
                ws.get_Range(celda, celda).Formula = isrCorriente;

                Console.WriteLine("[CONCEPTO]: I.S.R. CORRIENTE = " + isrCorriente);

                celda = celdaMes + "88"; //I.S.R. DIFERIDO
                ws.get_Range(celda, celda).Formula = isrDiferido;

                Console.WriteLine("[CONCEPTO]: I.S.R. DIFERIDO = " + isrDiferido);
            }
            else
            {
                celda = celdaMes + "46"; //GASTOS CORPORATIVOS
                ws.get_Range(celda, celda).Formula = gastoCorpo;

                Console.WriteLine("[CONCEPTO]: GASTOS CORPORATIVOS = " + gastoCorpo);

                celda = celdaMes + "64"; //I.S.R. CORRIENTE
                ws.get_Range(celda, celda).Formula = isrCorriente;

                Console.WriteLine("[CONCEPTO]: I.S.R. CORRIENTE = " + isrCorriente);

                celda = celdaMes + "82"; //I.S.R. DIFERIDO
                ws.get_Range(celda, celda).Formula = isrDiferido;

                Console.WriteLine("[CONCEPTO]: I.S.R. DIFERIDO = " + isrDiferido);
            }

            if (otroGastosProd == 0)
            {
                List<ConceptosContables> conceptosV1 = ConceptosContables.ListarRMV1(_db);
                List<ConceptosContables> conceptosV2 = ConceptosContables.ListarRMV2(_db);

                System.Data.DataTable dtRMWebV1 = proc.GetRMWebV1();
                System.Data.DataTable dtRMWebV1Acum = proc.GetRMWebV1Acumulado();
                System.Data.DataTable dtRMWebV2 = proc.GetRMWebV2();
                System.Data.DataTable dtRMWebV2Acum = proc.GetRMWebV2Acumulado();

                int r = 1;
                int vWeb = 0;
                int intAux = 0;
                DataRow[] drRMWebV1 = null;
                DataRow[] drRMWebV1Acum = null;
                DataRow[] drRMWebV2 = null;
                DataRow[] drRMWebV2Acum = null;

                foreach (ConceptosContables concepto in conceptosV2)
                {
                    Reporte rep = new Reporte();

                    r++;

                    rep.ID_CONCEPTO = concepto.Id;
                    rep.CONCEPTO = concepto.NombreConcepto;

                    if (idAgencia == 27)
                    {
                        drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 286)
                    {
                        drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 36)
                    {
                        drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 12)
                    {
                        drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 35)
                    {
                        drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 588)
                    {
                        drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 32)
                    {
                        drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 212)
                    {
                        drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else
                    {
                        drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                    }

                    if (idAgencia == 27)
                    {
                        drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 286)
                    {
                        drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",100) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 36)
                    {
                        drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",563,593,651) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 12)
                    {
                        drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 35)
                    {
                        drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",595,652) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 588)
                    {
                        drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",590) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 32)
                    {
                        drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",116) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else if (idAgencia == 212)
                    {
                        drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + idAgencia + ",33) AND ID_CONCEPTO = " + concepto.Id);
                    }
                    else
                    {
                        drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                        drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA = " + idAgencia + " AND ID_CONCEPTO = " + concepto.Id);
                    }

                    if (drRMWebV1.Length != 0)
                    {
                        vWeb = 0;

                        foreach (DataRow dr in drRMWebV1)
                        {
                            if ((concepto.Id == 1001) || (concepto.Id == 1003) || (concepto.Id == 1013) || (concepto.Id == 1015) || (concepto.Id == 1031)
                                 || (concepto.Id == 1032) || (concepto.Id == 1060) || (concepto.Id == 1064) || (concepto.Id == 1065))
                                vWeb += Convert.ToInt32(dr["VALOR"]);
                            else if (concepto.Id == 1033)
                            {
                                FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, idAgencia, FINA_EParametros.HorasFacturadasServVal);

                                if (resp1 != null && int.TryParse(resp1.Valor, out intAux))
                                {
                                    intAux = intAux == 0 ? 1 : intAux;

                                    vWeb += Convert.ToInt32(dr["VALOR"]) / intAux;
                                }
                            }
                            else
                                vWeb += Convert.ToInt32(dr["VALOR"]) / 1000;
                        }
                        rep.WEB_V1 = vWeb;

                        vWeb = 0;

                        foreach (DataRow dr in drRMWebV1Acum)
                        {
                            if ((concepto.Id == 1001) || (concepto.Id == 1003) || (concepto.Id == 1013) || (concepto.Id == 1015) || (concepto.Id == 1031)
                                 || (concepto.Id == 1032) || (concepto.Id == 1060) || (concepto.Id == 1064) || (concepto.Id == 1065))
                                vWeb += Convert.ToInt32(dr["VALOR"]);
                            else if (concepto.Id == 1033)
                            {
                                FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, idAgencia, FINA_EParametros.HorasFacturadasServVal);

                                if (resp1 != null && int.TryParse(resp1.Valor, out intAux))
                                {
                                    intAux = intAux == 0 ? 1 : intAux;

                                    vWeb += Convert.ToInt32(dr["VALOR"]) / intAux;
                                }
                            }
                            else
                                vWeb += Convert.ToInt32(dr["VALOR"]) / 1000;
                        }
                        rep.WEB_V1_ACUM = vWeb;
                    }

                    rep.DIFF_V1 = rep.EXCEL_V1 - rep.WEB_V1;
                    rep.DIFF_V1_ACUM = rep.EXCEL_V1_ACUM - rep.WEB_V1_ACUM;

                    if (drRMWebV2.Length != 0)
                    {
                        vWeb = 0;

                        foreach (DataRow dr in drRMWebV2)
                        {
                            if ((concepto.Id == 1001) || (concepto.Id == 1003) || (concepto.Id == 1013) || (concepto.Id == 1015) || (concepto.Id == 1031)
                                 || (concepto.Id == 1032) || (concepto.Id == 1060) || (concepto.Id == 1064) || (concepto.Id == 1065))
                                vWeb += Convert.ToInt32(dr["VALOR"]);
                            else if (concepto.Id == 1033)
                            {
                                FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, idAgencia, FINA_EParametros.HorasFacturadasServVal);

                                if (resp1 != null && int.TryParse(resp1.Valor, out intAux))
                                {
                                    intAux = intAux == 0 ? 1 : intAux;

                                    vWeb += Convert.ToInt32(dr["VALOR"]) / intAux;
                                }
                            }
                            else
                                vWeb += Convert.ToInt32(dr["VALOR"]) / 1000;
                        }
                        rep.WEB_V2 = vWeb;

                        vWeb = 0;

                        foreach (DataRow dr in drRMWebV2Acum)
                        {
                            if ((concepto.Id == 1001) || (concepto.Id == 1003) || (concepto.Id == 1013) || (concepto.Id == 1015) || (concepto.Id == 1031)
                                 || (concepto.Id == 1032) || (concepto.Id == 1060) || (concepto.Id == 1064) || (concepto.Id == 1065))
                                vWeb += Convert.ToInt32(dr["VALOR"]);
                            else if (concepto.Id == 1033)
                            {
                                FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, idAgencia, FINA_EParametros.HorasFacturadasServVal);

                                if (resp1 != null && int.TryParse(resp1.Valor, out intAux))
                                {
                                    intAux = intAux == 0 ? 1 : intAux;

                                    vWeb += Convert.ToInt32(dr["VALOR"]) / intAux;
                                }
                            }
                            else
                                vWeb += Convert.ToInt32(dr["VALOR"]) / 1000;
                        }
                        rep.WEB_V2_ACUM = vWeb;
                    }

                    rep.DIFF_V2 = rep.EXCEL_V2 - rep.WEB_V2;
                    rep.DIFF_V2_ACUM = rep.EXCEL_V2_ACUM - rep.WEB_V2_ACUM;

                    celda = GetMes();

                    if (version == "V1")
                    {
                        if (concepto.Id == 1049) //1049_OTROS (GASTOS) PRODUCTOS
                            celda += "24";
                    }
                    else
                    {
                        if (concepto.Id == 1049) //1049_OTROS (GASTOS) PRODUCTOS
                            celda += "24";
                    }

                    if (celda.Length >= 2)
                    {
                        if (version == "V1")
                            ws.get_Range(celda, celda).Formula = rep.WEB_V1;
                        else
                            ws.get_Range(celda, celda).Formula = rep.WEB_V2;
                    }

                    celda = "";

                    Console.WriteLine("[CONCEPTO]: " + concepto.Id + "_" + concepto.NombreConcepto);
                }
            }
        }

        public void LlenaPestaniaPX(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            int capturaPX = 0;

            //celda = "A1";

            //if (celda == "A1") //Razón Social
            //    ws.get_Range(celda, celda).Formula = razonSocial;

            //celda = "";

            List<ReportePX> lstPX = ReportePX.ListarPorIdAgenciaYAnio(_db, anio, idAgencia);

            foreach (ReportePX px in lstPX)
            {
                if (mes == 1)
                    capturaPX = capturaPX + Convert.ToInt32(px.Enero);
                else if (mes == 2)
                    capturaPX = capturaPX + Convert.ToInt32(px.Febrero);
                else if (mes == 3)
                    capturaPX = capturaPX + Convert.ToInt32(px.Marzo);
                else if (mes == 4)
                    capturaPX = capturaPX + Convert.ToInt32(px.Abril);
                else if (mes == 5)
                    capturaPX = capturaPX + Convert.ToInt32(px.Mayo);
                else if (mes == 6)
                    capturaPX = capturaPX + Convert.ToInt32(px.Junio);
                else if (mes == 7)
                    capturaPX = capturaPX + Convert.ToInt32(px.Julio);
                else if (mes == 8)
                    capturaPX = capturaPX + Convert.ToInt32(px.Agosto);
                else if (mes == 9)
                    capturaPX = capturaPX + Convert.ToInt32(px.Septiembre);
                else if (mes == 10)
                    capturaPX = capturaPX + Convert.ToInt32(px.Octubre);
                else if (mes == 11)
                    capturaPX = capturaPX + Convert.ToInt32(px.Noviembre);
                else if (mes == 12)
                    capturaPX = capturaPX + Convert.ToInt32(px.Diciembre);
            }

            celdaMes = GetMes();

            celda = celdaMes + "68"; //PARTIDA EXTRAORDINARIA
            ws.get_Range(celda, celda).Formula = capturaPX;

            Console.WriteLine("[CONCEPTO]: PARTIDA EXTRAORDINARIA = " + capturaPX);
        }

        public void LlenaPestaniaSI(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            int r = 0;
            
            List<SituacionInventario> lstSI = SituacionInventario.ListarPorAnioPorPeriodo(_db, idAgencia, anio, mes);

            celdaMes = GetMes();

            r = Convert.ToInt32(celdaMes);

            foreach (SituacionInventario si in lstSI)
            {
                celda = "B" + r; //AUTOS
                ws.get_Range(celda, celda).Formula = si.AUTOS;

                celda = "R" + r; //AUTOS
                ws.get_Range(celda, celda).Formula = si.AUTOS;

                Console.WriteLine("[CONCEPTO]: AUTOS = " + si.AUTOS);

                celda = "C" + r; //COMERCIALES
                ws.get_Range(celda, celda).Formula = si.COMERCIALES;

                celda = "S" + r; //COMERCIALES
                ws.get_Range(celda, celda).Formula = si.COMERCIALES;

                Console.WriteLine("[CONCEPTO]: COMERCIALES = " + si.COMERCIALES);

                celda = "D" + r; //AUTOFIN
                ws.get_Range(celda, celda).Formula = si.AUTOFIN;

                celda = "T" + r; //AUTOFIN
                ws.get_Range(celda, celda).Formula = si.AUTOFIN;

                Console.WriteLine("[CONCEPTO]: AUTOFIN = " + si.AUTOFIN);

                celda = "E" + r; //PLANTA
                ws.get_Range(celda, celda).Formula = si.PLANTA;

                celda = "U" + r; //PLANTA
                ws.get_Range(celda, celda).Formula = si.PLANTA;

                Console.WriteLine("[CONCEPTO]: PLANTA = " + si.PLANTA);

                celda = "F" + r; //PROPIAS_UNIDADES
                ws.get_Range(celda, celda).Formula = si.PROPIAS_UNIDADES;

                celda = "V" + r; //PROPIAS_UNIDADES
                ws.get_Range(celda, celda).Formula = si.PROPIAS_UNIDADES;

                Console.WriteLine("[CONCEPTO]: PROPIAS_UNIDADES = " + si.PROPIAS_UNIDADES);

                celda = "G" + r; //PROPIAS_IMPORTE
                ws.get_Range(celda, celda).Formula = si.PROPIAS_IMPORTE;

                celda = "W" + r; //PROPIAS_IMPORTE
                ws.get_Range(celda, celda).Formula = si.PROPIAS_IMPORTE;

                Console.WriteLine("[CONCEPTO]: PROPIAS_IMPORTE = " + si.PROPIAS_IMPORTE);

                celda = "H" + r; //FINANCIADASCI_UNIDADES
                ws.get_Range(celda, celda).Formula = si.FINANCIADASCI_UNIDADES;

                celda = "X" + r; //FINANCIADASCI_UNIDADES
                ws.get_Range(celda, celda).Formula = si.FINANCIADASCI_UNIDADES;

                Console.WriteLine("[CONCEPTO]: FINANCIADASCI_UNIDADES = " + si.FINANCIADASCI_UNIDADES);

                celda = "I" + r; //FINANCIADASCI_IMPORTE
                ws.get_Range(celda, celda).Formula = si.FINANCIADASCI_IMPORTE;

                celda = "Y" + r; //FINANCIADASCI_IMPORTE
                ws.get_Range(celda, celda).Formula = si.FINANCIADASCI_IMPORTE;

                Console.WriteLine("[CONCEPTO]: FINANCIADASCI_IMPORTE = " + si.FINANCIADASCI_IMPORTE);

                celda = "J" + r; //FINANCIADASSI_UNIDADES
                ws.get_Range(celda, celda).Formula = si.FINANCIADASSI_UNIDADES;

                celda = "Z" + r; //FINANCIADASSI_UNIDADES
                ws.get_Range(celda, celda).Formula = si.FINANCIADASSI_UNIDADES;

                Console.WriteLine("[CONCEPTO]: FINANCIADASSI_UNIDADES = " + si.FINANCIADASSI_UNIDADES);

                celda = "K" + r; //FINANCIADASSI_IMPORTE
                ws.get_Range(celda, celda).Formula = si.FINANCIADASSI_IMPORTE;

                celda = "AA" + r; //FINANCIADASSI_IMPORTE
                ws.get_Range(celda, celda).Formula = si.FINANCIADASSI_IMPORTE;

                Console.WriteLine("[CONCEPTO]: FINANCIADASSI_IMPORTE = " + si.FINANCIADASSI_IMPORTE);

                celda = "L" + r; //SEMINUEVOS_UNIDADES
                ws.get_Range(celda, celda).Formula = si.SEMINUEVOS_UNIDADES;

                Console.WriteLine("[CONCEPTO]: SEMINUEVOS_UNIDADES = " + si.SEMINUEVOS_UNIDADES);

                celda = "M" + r; //SEMINUEVOS_IMPORTE
                ws.get_Range(celda, celda).Formula = si.SEMINUEVOS_IMPORTE;

                Console.WriteLine("[CONCEPTO]: SEMINUEVOS_IMPORTE = " + si.SEMINUEVOS_IMPORTE);

                celda = "N" + r; //SERVICIO_IMPORTE
                ws.get_Range(celda, celda).Formula = si.SERVICIO_IMPORTE;

                celda = "AN" + r; //SERVICIO_IMPORTE
                ws.get_Range(celda, celda).Formula = si.SERVICIO_IMPORTE;

                Console.WriteLine("[CONCEPTO]: SERVICIO_IMPORTE = " + si.SERVICIO_IMPORTE);

                celda = "O" + r; //REFACCIONES_IMPORTE
                ws.get_Range(celda, celda).Formula = si.REFACCIONES_IMPORTE;

                celda = "AO" + r; //REFACCIONES_IMPORTE
                ws.get_Range(celda, celda).Formula = si.REFACCIONES_IMPORTE;

                Console.WriteLine("[CONCEPTO]: REFACCIONES_IMPORTE = " + si.REFACCIONES_IMPORTE);

                r++;
            }

            lstSI = SituacionInventario.ListarPorAnioPorPeriodoDetalleAU(_db, idAgencia, anio, mes);

            celdaMes = GetMes();

            r = Convert.ToInt32(celdaMes);

            foreach (SituacionInventario si in lstSI)
            {
                celda = "AC" + r; //AUTOS
                ws.get_Range(celda, celda).Formula = si.AUTOS;
                                
                Console.WriteLine("[CONCEPTO]: AU - AUTOS = " + si.AUTOS);

                celda = "AD" + r; //COMERCIALES
                ws.get_Range(celda, celda).Formula = si.COMERCIALES;
                
                Console.WriteLine("[CONCEPTO]: AU - COMERCIALES = " + si.COMERCIALES);

                celda = "AE" + r; //AUTOFIN
                ws.get_Range(celda, celda).Formula = si.AUTOFIN;
                                
                Console.WriteLine("[CONCEPTO]: AU - AUTOFIN = " + si.AUTOFIN);

                celda = "AF" + r; //PLANTA
                ws.get_Range(celda, celda).Formula = si.PLANTA;
                
                Console.WriteLine("[CONCEPTO]: AU - PLANTA = " + si.PLANTA);

                celda = "AG" + r; //PROPIAS_UNIDADES
                ws.get_Range(celda, celda).Formula = si.PROPIAS_UNIDADES;
                                
                Console.WriteLine("[CONCEPTO]: AU - PROPIAS_UNIDADES = " + si.PROPIAS_UNIDADES);

                celda = "AH" + r; //PROPIAS_IMPORTE
                ws.get_Range(celda, celda).Formula = si.PROPIAS_IMPORTE;
                                
                Console.WriteLine("[CONCEPTO]: AU - PROPIAS_IMPORTE = " + si.PROPIAS_IMPORTE);

                celda = "AI" + r; //FINANCIADASCI_UNIDADES
                ws.get_Range(celda, celda).Formula = si.FINANCIADASCI_UNIDADES;
                
                Console.WriteLine("[CONCEPTO]: AU - FINANCIADASCI_UNIDADES = " + si.FINANCIADASCI_UNIDADES);

                celda = "AJ" + r; //FINANCIADASCI_IMPORTE
                ws.get_Range(celda, celda).Formula = si.FINANCIADASCI_IMPORTE;
                                
                Console.WriteLine("[CONCEPTO]: AU - FINANCIADASCI_IMPORTE = " + si.FINANCIADASCI_IMPORTE);

                celda = "AK" + r; //FINANCIADASSI_UNIDADES
                ws.get_Range(celda, celda).Formula = si.FINANCIADASSI_UNIDADES;
                                
                Console.WriteLine("[CONCEPTO]: AU - FINANCIADASSI_UNIDADES = " + si.FINANCIADASSI_UNIDADES);

                celda = "AL" + r; //FINANCIADASSI_IMPORTE
                ws.get_Range(celda, celda).Formula = si.FINANCIADASSI_IMPORTE;
                                
                Console.WriteLine("[CONCEPTO]: AU - FINANCIADASSI_IMPORTE = " + si.FINANCIADASSI_IMPORTE);
                
                r++;
            }

            System.Data.DataTable dtSIREFA = new System.Data.DataTable();

            string query = "SELECT REDRETPG.FIREIDCIAU, REDRETPG.FIREANIO, REDRETPG.FIREMES, REDRETPG.FIREIDANTI, RECANTAB.FSREDESANT, REDRETPG.FIRETOTPA, ANCTIPGO.FSANTPDES, REDRETPG.FNREMONDPA\r\n" +
                "FROM [PREFIX]REFA.REDRETPG REDRETPG\r\n" +
                "INNER JOIN [PREFIX]REFA.RECANTAB RECANTAB\r\nON REDRETPG.FIREIDANTI = RECANTAB.FIREIDANTI AND REDRETPG.FIRESTATUS = RECANTAB.FIRESTATUS\r\n" +
                "INNER JOIN [PREFIX]AUT.ANCTIPGO ANCTIPGO\r\nON REDRETPG.FIRETOTPA = ANCTIPGO.FNANTPIDE AND REDRETPG.FIRESTATUS = ANCTIPGO.FIANSTATU\r\n" +
                "WHERE \r\nREDRETPG.FIRESTATUS = 1\r\n" +
                "AND REDRETPG.FIREANIO = " + anio + "\r\n" +
                "AND REDRETPG.FIREMES = " + mes + "\r\n" +
                "AND REDRETPG.FIREIDCIAU = " + idAgencia + "\r\n" +
                "ORDER BY REDRETPG.FIREIDANTI";

            dtSIREFA = _db.GetDataTable(query);

            r = Convert.ToInt32(celdaMes);

            int idTipoPago = 0;
            double importe = 0;

            foreach (DataRow dr in dtSIREFA.Rows)
            {
                idTipoPago = Convert.ToInt32(dr["FIRETOTPA"]);
                importe = Convert.ToDouble(dr["FNREMONDPA"]);

                if (idTipoPago == 2) //2_PROPIO
                {
                    celda = "AP" + r; //REFACCIONES_IMPORTE_PROPIAS
                    ws.get_Range(celda, celda).Formula = importe;

                    Console.WriteLine("[CONCEPTO]: REFACCIONES_IMPORTE_PROPIAS = " + importe);
                }
                else if (idTipoPago == 5) //5_FINANCIADO CON INTERESES
                {
                    celda = "AQ" + r; //REFACCIONES_IMPORTE_FINANCIADO CON INTERESES
                    ws.get_Range(celda, celda).Formula = importe;

                    Console.WriteLine("[CONCEPTO]: REFACCIONES_IMPORTE_FINANCIADO CON INTERESES = " + importe);
                }
                else if (idTipoPago == 4) //4_FINANCIADO SIN INTERESES
                {
                    celda = "AR" + r; //REFACCIONES_IMPORTE_FINANCIADO SIN INTERESES
                    ws.get_Range(celda, celda).Formula = importe;

                    Console.WriteLine("[CONCEPTO]: REFACCIONES_IMPORTE_FINANCIADO SIN INTERESES = " + importe);
                }

                r++;
            }
        }

        public void LlenaPestaniaSC(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            int r = 0;

            List<SituacionCartera> lstSC = SituacionCartera.ListarPorPeriodoSC(_db, anio, mes, idAgencia);

            celdaMes = GetMes();
                        
            foreach (SituacionCartera sc in lstSC)
            {
                if (sc.IdRango != 1)
                {
                    r = Convert.ToInt32(celdaMes) + sc.IdRango;

                    celda = "C" + r; //DocsMesAnterior
                    ws.get_Range(celda, celda).Formula = sc.DocsMesAnterior;

                    Console.WriteLine("[CONCEPTO]: DocsMesAnterior = " + sc.DocsMesAnterior);

                    celda = "D" + r; //DocsMesActual
                    ws.get_Range(celda, celda).Formula = sc.DocsMesActual;

                    Console.WriteLine("[CONCEPTO]: DocsMesActual = " + sc.DocsMesActual);

                    celda = "E" + r; //AutosNuevMesAnterior
                    ws.get_Range(celda, celda).Formula = sc.AutosNuevMesAnterior;

                    Console.WriteLine("[CONCEPTO]: AutosNuevMesAnterior = " + sc.AutosNuevMesAnterior);

                    celda = "F" + r; //AutosNuevMesActual
                    ws.get_Range(celda, celda).Formula = sc.AutosNuevMesActual;

                    Console.WriteLine("[CONCEPTO]: AutosNuevMesActual = " + sc.AutosNuevMesActual);

                    celda = "G" + r; //AutoFinMesAnterior
                    ws.get_Range(celda, celda).Formula = sc.AutoFinMesAnterior;

                    Console.WriteLine("[CONCEPTO]: AutoFinMesAnterior = " + sc.AutoFinMesAnterior);

                    celda = "H" + r; //AutoFinMesActual
                    ws.get_Range(celda, celda).Formula = sc.AutoFinMesActual;

                    Console.WriteLine("[CONCEPTO]: AutoFinMesActual = " + sc.AutoFinMesActual);

                    celda = "I" + r; //ServMesAnterior
                    ws.get_Range(celda, celda).Formula = sc.ServMesAnterior;

                    Console.WriteLine("[CONCEPTO]: ServMesAnterior = " + sc.ServMesAnterior);

                    celda = "J" + r; //ServMesActual
                    ws.get_Range(celda, celda).Formula = sc.ServMesActual;

                    Console.WriteLine("[CONCEPTO]: ServMesActual = " + sc.ServMesActual);

                    celda = "K" + r; //RefacMesAnterior
                    ws.get_Range(celda, celda).Formula = sc.RefacMesAnterior;

                    Console.WriteLine("[CONCEPTO]: RefacMesAnterior = " + sc.RefacMesAnterior);

                    celda = "L" + r; //RefacMesActual
                    ws.get_Range(celda, celda).Formula = sc.RefacMesActual;

                    Console.WriteLine("[CONCEPTO]: RefacMesActual = " + sc.RefacMesActual);

                    celda = "M" + r; //PlantaMesAnterior
                    ws.get_Range(celda, celda).Formula = sc.PlantaMesAnterior;

                    Console.WriteLine("[CONCEPTO]: PlantaMesAnterior = " + sc.PlantaMesAnterior);

                    celda = "N" + r; //PlantaMesActual
                    ws.get_Range(celda, celda).Formula = sc.PlantaMesActual;

                    Console.WriteLine("[CONCEPTO]: PlantaMesActual = " + sc.PlantaMesActual);

                    //r++;
                }
            }
        }

        public void LlenaPestaniaGANMT(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<DetalleSaldo> lstDetSaldo = DetalleSaldo.ListarGeneral(_db, mes, anio, 1, GetAgenciasYSucursales(_db, idAgencia)); //1_ANT_AUTOS_NUEVOS

            celdaMes = GetMes();
                        
            foreach (DetalleSaldo det in lstDetSaldo)
            {
                if (det.IdConcepto == 291) //291_Sueldos y Honorarios Operativos
                    celda = celdaMes + "4";
                else if (det.IdConcepto == 292) //292_Comisiones Operativos
                    celda = celdaMes + "5";
                else if (det.IdConcepto == 293) //293_Sueldos Gerentes Operativos
                    celda = celdaMes + "6";
                else if (det.IdConcepto == 294) //294_Gratificación Anual Operativos
                    celda = celdaMes + "7";
                else if (det.IdConcepto == 295) //295_Sueldos y Honorarios Administración
                    celda = celdaMes + "8";
                else if (det.IdConcepto == 296) //296_Gratificación Anual Administración
                    celda = celdaMes + "9";
                else if (det.IdConcepto == 297) //297_Previsión Social
                    celda = celdaMes + "10";
                else if (det.IdConcepto == 298) //298_Impuestos Laboral
                    celda = celdaMes + "11";
                else if (det.IdConcepto == 299) //299_Traslado de Unidades
                    celda = celdaMes + "14";
                else if (det.IdConcepto == 300) //300_Fletes y Embarques
                    celda = celdaMes + "15";
                else if (det.IdConcepto == 301) //301_Combustibles y Lubricantes
                    celda = celdaMes + "16";
                else if (det.IdConcepto == 302) //302_Acondicionamiento Unidades
                    celda = celdaMes + "17";
                else if (det.IdConcepto == 303) //303_Garantías / Cortesías
                    celda = celdaMes + "18";
                else if (det.IdConcepto == 304) //304_Cuotas y Suscripciones
                    celda = celdaMes + "19";
                else if (det.IdConcepto == 305) //305_Teléfono y Correo
                    celda = celdaMes + "20";
                else if (det.IdConcepto == 306) //306_Publicidad y Promoción
                    celda = celdaMes + "21";
                else if (det.IdConcepto == 307) //307_Herramientas de Taller
                    celda = celdaMes + "22";
                else if (det.IdConcepto == 308) //308_Manto. de Unidades
                    celda = celdaMes + "23";
                else if (det.IdConcepto == 309) //309_Manto. de Equipo
                    celda = celdaMes + "24";
                else if (det.IdConcepto == 310) //310_Materiales Diversos Taller
                    celda = celdaMes + "25";
                else if (det.IdConcepto == 311) //311_Asesoria Externos
                    celda = celdaMes + "26";
                else if (det.IdConcepto == 312) //312_Capacitación
                    celda = celdaMes + "27";
                else if (det.IdConcepto == 313) //313_Papelería
                    celda = celdaMes + "28";
                else if (det.IdConcepto == 314) //314_Gastos de Viaje
                    celda = celdaMes + "29";
                else if (det.IdConcepto == 315) //315_Traslado de Valores
                    celda = celdaMes + "30";
                else if (det.IdConcepto == 316) //316_No Deducibles
                    celda = celdaMes + "31";
                else if (det.IdConcepto == 317) //317_Vigilancia y Aseo
                    celda = celdaMes + "32";
                else if (det.IdConcepto == 318) //318_Luz y Agua
                    celda = celdaMes + "33";
                else if (det.IdConcepto == 319) //319_Impuestos y Derechos
                    celda = celdaMes + "34";
                else if (det.IdConcepto == 320) //320_Depreciación
                    celda = celdaMes + "37";
                else if (det.IdConcepto == 321) //321_Arrendamiento Inmuebles
                    celda = celdaMes + "38";
                else if (det.IdConcepto == 322) //322_Seguros y Fianzas
                    celda = celdaMes + "39";
                else if (det.IdConcepto == 323) //323_Publicidad y Promoción
                    celda = celdaMes + "40";
                else if (det.IdConcepto == 324) //324_Asesoria y Honorarios
                    celda = celdaMes + "41";
                else if (det.IdConcepto == 325) //325_Arrendamiento Inmuebles
                    celda = celdaMes + "42";
                else if (det.IdConcepto == 326) //326_Manto. de Edificio
                    celda = celdaMes + "43";

                if (celda.Length >= 2)
                {
                    ws.get_Range(celda, celda).Formula = det.Actual;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + det.Cuenta + " = " + det.Actual);
            }
        }

        public void LlenaPestaniaGANMF(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<DetalleSaldo> lstDetSaldo = DetalleSaldo.ListarGeneral(_db, mes, anio, 19, GetAgenciasYSucursales(_db, idAgencia)); //19_ANT_AUTOS_NUEVOS_FLOTILLAS

            celdaMes = GetMes();

            foreach (DetalleSaldo det in lstDetSaldo)
            {
                if (det.IdConcepto == 365) //365_Sueldos y Honorarios Operativos
                    celda = celdaMes + "4";
                else if (det.IdConcepto == 366) //366_Comisiones Departamentos Operativos
                    celda = celdaMes + "5";
                else if (det.IdConcepto == 367) //367_Sueldos Gerentes Operativos
                    celda = celdaMes + "6";
                else if (det.IdConcepto == 368) //368_Gratificación Anual Operativos
                    celda = celdaMes + "7";
                else if (det.IdConcepto == 369) //369_Sueldos y Honorarios Administración
                    celda = celdaMes + "8";
                else if (det.IdConcepto == 370) //370_Gratificación Anual Administración
                    celda = celdaMes + "9";
                else if (det.IdConcepto == 371) //371_Previsión Social
                    celda = celdaMes + "10";
                else if (det.IdConcepto == 372) //372_Impuestos Laboral
                    celda = celdaMes + "11";
                else if (det.IdConcepto == 373) //373_Traslado de Unidades
                    celda = celdaMes + "14";
                else if (det.IdConcepto == 374) //374_Fletes y Embarques
                    celda = celdaMes + "15";
                else if (det.IdConcepto == 375) //375_Combustibles y Lubricantes
                    celda = celdaMes + "16";
                else if (det.IdConcepto == 376) //376_Acondicionamiento Unidades
                    celda = celdaMes + "17";
                else if (det.IdConcepto == 377) //377_Garantías / Cortesías
                    celda = celdaMes + "18";
                else if (det.IdConcepto == 378) //378_Cuotas y Suscripciones
                    celda = celdaMes + "19";
                else if (det.IdConcepto == 379) //379_Teléfono y Correo
                    celda = celdaMes + "20";
                else if (det.IdConcepto == 380) //380_Publicidad y Promoción
                    celda = celdaMes + "21";
                else if (det.IdConcepto == 381) //381_Herramientas de Taller
                    celda = celdaMes + "22";
                else if (det.IdConcepto == 382) //382_Manto. de Unidades
                    celda = celdaMes + "23";
                else if (det.IdConcepto == 383) //383_Manto. de Equipo
                    celda = celdaMes + "24";
                else if (det.IdConcepto == 384) //384_Materiales Diversos Taller
                    celda = celdaMes + "25";
                else if (det.IdConcepto == 385) //385_Asesoria Externos
                    celda = celdaMes + "26";
                else if (det.IdConcepto == 386) //386_Capacitación
                    celda = celdaMes + "27";
                else if (det.IdConcepto == 387) //387_Papelería
                    celda = celdaMes + "28";
                else if (det.IdConcepto == 388) //388_Gastos de Viaje
                    celda = celdaMes + "29";
                else if (det.IdConcepto == 389) //389_Traslado de Valores
                    celda = celdaMes + "30";
                else if (det.IdConcepto == 390) //390_No Deducibles
                    celda = celdaMes + "31";
                else if (det.IdConcepto == 391) //391_Vigilancia y Aseo
                    celda = celdaMes + "32";
                else if (det.IdConcepto == 392) //392_Luz y Agua
                    celda = celdaMes + "33";
                else if (det.IdConcepto == 393) //393_Impuestos y Derechos
                    celda = celdaMes + "34";
                else if (det.IdConcepto == 394) //394_Depreciación
                    celda = celdaMes + "37";
                else if (det.IdConcepto == 395) //395_Arrendamiento Inmuebles
                    celda = celdaMes + "38";
                else if (det.IdConcepto == 396) //396_Seguros y Fianzas
                    celda = celdaMes + "39";
                else if (det.IdConcepto == 397) //397_Publicidad y Promoción
                    celda = celdaMes + "40";
                else if (det.IdConcepto == 398) //398_Asesoria y Honorarios
                    celda = celdaMes + "41";
                else if (det.IdConcepto == 399) //399_Arrendamiento Inmuebles
                    celda = celdaMes + "42";
                else if (det.IdConcepto == 400) //400_Manto. de Edificio
                    celda = celdaMes + "43";
                else if (det.IdConcepto == 401) //401_Mejoras al Inmueble
                    celda = celdaMes + "44";

                if (celda.Length >= 2)
                {
                    ws.get_Range(celda, celda).Formula = det.Actual;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + det.Cuenta + " = " + det.Actual);
            }
        }

        public void LlenaPestaniaGANMVI(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<DetalleSaldo> lstDetSaldo = DetalleSaldo.ListarGeneral(_db, mes, anio, 21, GetAgenciasYSucursales(_db, idAgencia)); //21_ANT_AUTOS_NUEVOS_DIGITAL

            celdaMes = GetMes();

            foreach (DetalleSaldo det in lstDetSaldo)
            {
                if (det.IdConcepto == 402) //402_Sueldos y Honorarios Operativos
                    celda = celdaMes + "4";
                else if (det.IdConcepto == 403) //403_Comisiones Operativos
                    celda = celdaMes + "5";
                else if (det.IdConcepto == 404) //404_Sueldos Gerentes Operativos
                    celda = celdaMes + "6";
                else if (det.IdConcepto == 405) //405_Gratificación Anual Operativos
                    celda = celdaMes + "7";
                else if (det.IdConcepto == 406) //406_Sueldos y Honorarios Administración
                    celda = celdaMes + "8";
                else if (det.IdConcepto == 407) //407_Gratificación Anual Administración
                    celda = celdaMes + "9";
                else if (det.IdConcepto == 408) //408_Previsión Social
                    celda = celdaMes + "10";
                else if (det.IdConcepto == 409) //409_Impuestos Laboral
                    celda = celdaMes + "11";
                else if (det.IdConcepto == 410) //410_Traslado de Unidades
                    celda = celdaMes + "14";
                else if (det.IdConcepto == 411) //411_Fletes y Embarques
                    celda = celdaMes + "15";
                else if (det.IdConcepto == 412) //412_Combustibles y Lubricantes
                    celda = celdaMes + "16";
                else if (det.IdConcepto == 413) //413_Acondicionamiento Unidades
                    celda = celdaMes + "17";
                else if (det.IdConcepto == 414) //414_Garantías / Cortesías
                    celda = celdaMes + "18";
                else if (det.IdConcepto == 415) //415_Cuotas y Suscripciones
                    celda = celdaMes + "19";
                else if (det.IdConcepto == 416) //416_Teléfono y Correo
                    celda = celdaMes + "20";
                else if (det.IdConcepto == 417) //417_Publicidad y Promoción
                    celda = celdaMes + "21";
                else if (det.IdConcepto == 418) //418_Herramientas de Taller
                    celda = celdaMes + "22";
                else if (det.IdConcepto == 419) //419_Manto. de Unidades
                    celda = celdaMes + "23";
                else if (det.IdConcepto == 420) //420_Manto. de Equipo
                    celda = celdaMes + "24";
                else if (det.IdConcepto == 421) //421_Materiales Diversos Taller
                    celda = celdaMes + "25";
                else if (det.IdConcepto == 422) //422_Asesoria Externos
                    celda = celdaMes + "26";
                else if (det.IdConcepto == 423) //423_Capacitación
                    celda = celdaMes + "27";
                else if (det.IdConcepto == 424) //424_Papelería
                    celda = celdaMes + "28";
                else if (det.IdConcepto == 425) //425_Gastos de Viaje
                    celda = celdaMes + "29";
                else if (det.IdConcepto == 426) //426_Traslado de Valores
                    celda = celdaMes + "30";
                else if (det.IdConcepto == 427) //427_No Deducibles
                    celda = celdaMes + "31";
                else if (det.IdConcepto == 428) //428_Vigilancia y Aseo
                    celda = celdaMes + "32";
                else if (det.IdConcepto == 429) //429_Luz y Agua
                    celda = celdaMes + "33";
                else if (det.IdConcepto == 430) //430_Impuestos y Derechos
                    celda = celdaMes + "34";
                else if (det.IdConcepto == 431) //431_Depreciación
                    celda = celdaMes + "37";
                else if (det.IdConcepto == 432) //432_Arrendamiento Inmuebles
                    celda = celdaMes + "38";
                else if (det.IdConcepto == 433) //433_Seguros y Fianzas
                    celda = celdaMes + "39";
                else if (det.IdConcepto == 434) //434_Publicidad y Promoción
                    celda = celdaMes + "40";
                else if (det.IdConcepto == 435) //435_Asesoria y Honorarios
                    celda = celdaMes + "41";
                else if (det.IdConcepto == 436) //436_Arrendamiento Inmuebles
                    celda = celdaMes + "42";
                else if (det.IdConcepto == 437) //437_Manto. de Edificio
                    celda = celdaMes + "43";
                else if (det.IdConcepto == 438) //438_Mejoras al Inmueble
                    celda = celdaMes + "44";

                if (celda.Length >= 2)
                {
                    ws.get_Range(celda, celda).Formula = det.Actual;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + det.Cuenta + " = " + det.Actual);
            }
        }

        public void LlenaPestaniaGGNM(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<DetalleSaldo> lstDetSaldo = DetalleSaldo.ListarGeneral(_db, mes, anio, 13, GetAgenciasYSucursales(_db, idAgencia)); //13_GDN_GERENCIA_DE_NEGOCIOS

            celdaMes = GetMes();

            foreach (DetalleSaldo det in lstDetSaldo)
            {
                if (det.IdConcepto == 439) //439_Sueldos y Honorarios Operativos
                    celda = celdaMes + "4";
                else if (det.IdConcepto == 440) //440_Comisiones Operativos
                    celda = celdaMes + "5";
                else if (det.IdConcepto == 441) //441_Sueldos Gerentes Operativos
                    celda = celdaMes + "6";
                else if (det.IdConcepto == 442) //442_Gratificación Anual Operativos
                    celda = celdaMes + "7";
                else if (det.IdConcepto == 443) //443_Sueldos y Honorarios Administración
                    celda = celdaMes + "8";
                else if (det.IdConcepto == 444) //444_Gratificación Anual Administración
                    celda = celdaMes + "9";
                else if (det.IdConcepto == 445) //445_Previsión Social
                    celda = celdaMes + "10";
                else if (det.IdConcepto == 446) //446_Impuestos Laboral
                    celda = celdaMes + "11";
                else if (det.IdConcepto == 447) //447_Traslado de Unidades por Terceros
                    celda = celdaMes + "14";
                else if (det.IdConcepto == 448) //448_Gastos por Fletes y Embarques
                    celda = celdaMes + "15";
                else if (det.IdConcepto == 449) //449_Combustibles y Lubricantes
                    celda = celdaMes + "16";
                else if (det.IdConcepto == 450) //450_Acondicionamiento Unidades
                    celda = celdaMes + "17";
                else if (det.IdConcepto == 451) //451_Garantías / Cortesías
                    celda = celdaMes + "18";
                else if (det.IdConcepto == 452) //452_Cuotas y Suscripciones
                    celda = celdaMes + "19";
                else if (det.IdConcepto == 453) //453_Teléfono y Correo
                    celda = celdaMes + "20";
                else if (det.IdConcepto == 454) //454_Publicidad y Promoción
                    celda = celdaMes + "21";
                else if (det.IdConcepto == 455) //455_Herramientas de Taller
                    celda = celdaMes + "22";
                else if (det.IdConcepto == 456) //456_Mantenimiento De Unidades
                    celda = celdaMes + "23";
                else if (det.IdConcepto == 457) //457_Manto. de Equipo
                    celda = celdaMes + "24";
                else if (det.IdConcepto == 458) //458_Materiales Diversos Del Taller
                    celda = celdaMes + "25";
                else if (det.IdConcepto == 459) //459_Asesoria Externos
                    celda = celdaMes + "26";
                else if (det.IdConcepto == 460) //460_Capacitación
                    celda = celdaMes + "27";
                else if (det.IdConcepto == 461) //461_Papelería
                    celda = celdaMes + "28";
                else if (det.IdConcepto == 462) //462_Gastos de Viaje
                    celda = celdaMes + "29";
                else if (det.IdConcepto == 463) //463_Traslado de Valores
                    celda = celdaMes + "30";
                else if (det.IdConcepto == 464) //464_No Deducibles
                    celda = celdaMes + "31";
                else if (det.IdConcepto == 465) //465_Vigilancia y Aseo
                    celda = celdaMes + "32";
                else if (det.IdConcepto == 466) //466_Luz y Agua
                    celda = celdaMes + "33";
                else if (det.IdConcepto == 467) //467_Impuestos y Derechos
                    celda = celdaMes + "34";
                else if (det.IdConcepto == 468) //468_Depreciación
                    celda = celdaMes + "37";
                else if (det.IdConcepto == 469) //469_Arrendamiento Inmuebles
                    celda = celdaMes + "38";
                else if (det.IdConcepto == 470) //470_Seguros y Fianzas
                    celda = celdaMes + "39";
                else if (det.IdConcepto == 471) //471_Publicidad Y Promocion Corporativa
                    celda = celdaMes + "40";
                else if (det.IdConcepto == 472) //472_Asesoria y Honorarios Corporativos
                    celda = celdaMes + "41";
                else if (det.IdConcepto == 473) //473_Arrendamiento Inmuebles
                    celda = celdaMes + "42";
                else if (det.IdConcepto == 474) //474_Manto. de Edificio
                    celda = celdaMes + "43";

                if (celda.Length >= 2)
                {
                    ws.get_Range(celda, celda).Formula = det.Actual;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + det.Cuenta + " = " + det.Actual);
            }
        }

        public void LlenaPestaniaGASM(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<DetalleSaldo> lstDetSaldo = DetalleSaldo.ListarGeneral(_db, mes, anio, 4, GetAgenciasYSucursales(_db, idAgencia)); //4_ASN_AUTOS_SEMI_NUEVOS

            celdaMes = GetMes();

            foreach (DetalleSaldo det in lstDetSaldo)
            {
                if (det.IdConcepto == 476) //476_Sueldos y Honorarios Deptos. Operativos
                    celda = celdaMes + "4";
                else if (det.IdConcepto == 477) //477_Comisiones Departamentos Operativos
                    celda = celdaMes + "5";
                else if (det.IdConcepto == 478) //478_Sueldos,  Honorarios, Com. y Prest. Gerentes
                    celda = celdaMes + "6";
                else if (det.IdConcepto == 479) //479_Gratificación Anual Deptos. Operativos
                    celda = celdaMes + "7";
                else if (det.IdConcepto == 480) //480_Sueldos y Honorarios  Depto. Admón.
                    celda = celdaMes + "8";
                else if (det.IdConcepto == 481) //481_Gratificación Anual Depto. Admon
                    celda = celdaMes + "9";
                else if (det.IdConcepto == 482) //482_Previsión Social
                    celda = celdaMes + "10";
                else if (det.IdConcepto == 483) //483_Impuestos Derivados de la Relación Laboral
                    celda = celdaMes + "11";
                else if (det.IdConcepto == 484) //484_Traslado de Unidades por Terceros
                    celda = celdaMes + "14";
                else if (det.IdConcepto == 485) //485_Gastos por Fletes y Embarques
                    celda = celdaMes + "15";
                else if (det.IdConcepto == 486) //486_Combustibles Y Lubricantes
                    celda = celdaMes + "16";
                else if (det.IdConcepto == 487) //487_Acondicionamiento Unidades Nuevas
                    celda = celdaMes + "17";
                else if (det.IdConcepto == 488) //488_Garantías / Cortesías
                    celda = celdaMes + "18";
                else if (det.IdConcepto == 489) //489_Cuotas y Suscripciones
                    celda = celdaMes + "19";
                else if (det.IdConcepto == 490) //490_Teléfono y Correo
                    celda = celdaMes + "20";
                else if (det.IdConcepto == 491) //491_Publicidad y Promoción
                    celda = celdaMes + "21";
                else if (det.IdConcepto == 492) //492_Herramientas del Taller
                    celda = celdaMes + "22";
                else if (det.IdConcepto == 493) //493_Mantenimiento de Unidades
                    celda = celdaMes + "23";
                else if (det.IdConcepto == 494) //494_Mantenimiento de Equipo
                    celda = celdaMes + "24";
                else if (det.IdConcepto == 495) //495_Materiales Diversos del Taller
                    celda = celdaMes + "25";
                else if (det.IdConcepto == 496) //496_Asesoria Externos
                    celda = celdaMes + "26";
                else if (det.IdConcepto == 497) //497_Capacitación y Adiestramiento
                    celda = celdaMes + "27";
                else if (det.IdConcepto == 498) //498_Papelería y Artículos de Escritorio
                    celda = celdaMes + "28";
                else if (det.IdConcepto == 499) //499_Gastos de Viaje
                    celda = celdaMes + "29";
                else if (det.IdConcepto == 500) //500_Traslado de Valores
                    celda = celdaMes + "30";
                else if (det.IdConcepto == 501) //501_No Deducibles
                    celda = celdaMes + "31";
                else if (det.IdConcepto == 502) //502_Vigilancia y Aseo
                    celda = celdaMes + "32";
                else if (det.IdConcepto == 503) //503_Luz y Agua
                    celda = celdaMes + "33";
                else if (det.IdConcepto == 504) //504_Otros Impuestos y Derechos
                    celda = celdaMes + "34";
                else if (det.IdConcepto == 505) //505_Depreciación y Amortización
                    celda = celdaMes + "37";
                else if (det.IdConcepto == 506) //506_Arrendamiento de Inmuebles
                    celda = celdaMes + "38";
                else if (det.IdConcepto == 507) //507_Seguros y Fianzas
                    celda = celdaMes + "39";
                else if (det.IdConcepto == 508) //508_Publicidad y Promoción Corporativa
                    celda = celdaMes + "40";
                else if (det.IdConcepto == 509) //509_Asesoria y Honorarios
                    celda = celdaMes + "41";
                else if (det.IdConcepto == 510) //510_Arrendamiento de Inmuebles Corporativos
                    celda = celdaMes + "42";
                else if (det.IdConcepto == 511) //511_Mantenimiento de Edificio
                    celda = celdaMes + "43";

                if (celda.Length >= 2)
                {
                    ws.get_Range(celda, celda).Formula = det.Actual;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + det.Cuenta + " = " + det.Actual);
            }
        }

        public void LlenaPestaniaGSM(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<DetalleSaldo> lstDetSaldo = DetalleSaldo.ListarGeneral(_db, mes, anio, 5, GetAgenciasYSucursales(_db, idAgencia)); //5_SER_SERVICIO

            celdaMes = GetMes();

            foreach (DetalleSaldo det in lstDetSaldo)
            {
                if (det.IdConcepto == 513) //513_Sueldos y Honorarios Deptos. Operativos
                    celda = celdaMes + "4";
                else if (det.IdConcepto == 514) //514_Comisiones Departamentos Operativos
                    celda = celdaMes + "5";
                else if (det.IdConcepto == 515) //515_Sueldos,  Honorarios, Com. y Prest. Gerentes
                    celda = celdaMes + "6";
                else if (det.IdConcepto == 516) //516_Gratificación Anual Deptos. Operativos
                    celda = celdaMes + "7";
                else if (det.IdConcepto == 517) //517_Sueldos y Honorarios  Depto. Admón.
                    celda = celdaMes + "8";
                else if (det.IdConcepto == 518) //518_Gratificación Anual Depto. Admon
                    celda = celdaMes + "9";
                else if (det.IdConcepto == 519) //519_Previsión Social
                    celda = celdaMes + "10";
                else if (det.IdConcepto == 520) //520_Impuestos Derivados de la Relación Laboral
                    celda = celdaMes + "11";
                else if (det.IdConcepto == 521) //521_Traslado de Unidades por Terceros
                    celda = celdaMes + "14";
                else if (det.IdConcepto == 522) //522_Gastos por Fletes y Embarques
                    celda = celdaMes + "15";
                else if (det.IdConcepto == 523) //523_Combustibles Y Lubricantes
                    celda = celdaMes + "16";
                else if (det.IdConcepto == 524) //524_Acondicionamiento Unidades Nuevas
                    celda = celdaMes + "17";
                else if (det.IdConcepto == 525) //525_Garantías / Cortesías
                    celda = celdaMes + "18";
                else if (det.IdConcepto == 526) //526_Cuotas y Suscripciones
                    celda = celdaMes + "19";
                else if (det.IdConcepto == 527) //527_Teléfono y Correo
                    celda = celdaMes + "20";
                else if (det.IdConcepto == 528) //528_Publicidad y Promoción
                    celda = celdaMes + "21";
                else if (det.IdConcepto == 529) //529_Herramientas del Taller
                    celda = celdaMes + "22";
                else if (det.IdConcepto == 530) //530_Mantenimiento de Unidades
                    celda = celdaMes + "23";
                else if (det.IdConcepto == 531) //531_Mantenimiento de Equipo
                    celda = celdaMes + "24";
                else if (det.IdConcepto == 532) //532_Materiales Diversos del Taller
                    celda = celdaMes + "25";
                else if (det.IdConcepto == 533) //533_Asesoria Externos
                    celda = celdaMes + "26";
                else if (det.IdConcepto == 534) //534_Capacitación y Adiestramiento
                    celda = celdaMes + "27";
                else if (det.IdConcepto == 535) //535_Papelería y Artículos de Escritorio
                    celda = celdaMes + "28";
                else if (det.IdConcepto == 536) //536_Gastos de Viaje
                    celda = celdaMes + "29";
                else if (det.IdConcepto == 537) //537_Traslado de Valores
                    celda = celdaMes + "30";
                else if (det.IdConcepto == 538) //538_No Deducibles
                    celda = celdaMes + "31";
                else if (det.IdConcepto == 539) //539_Vigilancia y Aseo
                    celda = celdaMes + "32";
                else if (det.IdConcepto == 540) //540_Luz y Agua
                    celda = celdaMes + "33";
                else if (det.IdConcepto == 541) //541_Otros Impuestos y Derechos
                    celda = celdaMes + "34";
                else if (det.IdConcepto == 542) //542_Depreciación y Amortización
                    celda = celdaMes + "37";
                else if (det.IdConcepto == 543) //543_Arrendamiento de Inmuebles
                    celda = celdaMes + "38";
                else if (det.IdConcepto == 544) //544_Seguros y Fianzas
                    celda = celdaMes + "39";
                else if (det.IdConcepto == 545) //545_Publicidad y Promoción Corporativa
                    celda = celdaMes + "40";
                else if (det.IdConcepto == 546) //546_Asesoria y Honorarios
                    celda = celdaMes + "41";
                else if (det.IdConcepto == 547) //547_Arrendamiento de Inmuebles Corporativos
                    celda = celdaMes + "42";
                else if (det.IdConcepto == 548) //548_Mantenimiento de Edificio
                    celda = celdaMes + "43";
                else if (det.IdConcepto == 549) //549_Mejoras al Inmueble
                    celda = celdaMes + "44";

                if (celda.Length >= 2)
                {
                    ws.get_Range(celda, celda).Formula = det.Actual;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + det.Cuenta + " = " + det.Actual);
            }
        }

        public void LlenaPestaniaGHPM(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<DetalleSaldo> lstDetSaldo = DetalleSaldo.ListarGeneral(_db, mes, anio, 6, GetAgenciasYSucursales(_db, idAgencia)); //6_HYP_HYP

            celdaMes = GetMes();

            foreach (DetalleSaldo det in lstDetSaldo)
            {
                if (det.IdConcepto == 550) //550_Sueldos y Honorarios Deptos. Operativos
                    celda = celdaMes + "4";
                else if (det.IdConcepto == 551) //551_Comisiones Departamentos Operativos
                    celda = celdaMes + "5";
                else if (det.IdConcepto == 552) //552_Sueldos,  Honorarios, Com. y Prest. Gerentes
                    celda = celdaMes + "6";
                else if (det.IdConcepto == 553) //553_Gratificación Anual Deptos. Operativos
                    celda = celdaMes + "7";
                else if (det.IdConcepto == 554) //554_Sueldos y Honorarios  Depto. Admón.
                    celda = celdaMes + "8";
                else if (det.IdConcepto == 555) //555_Gratificación Anual Depto. Admon
                    celda = celdaMes + "9";
                else if (det.IdConcepto == 556) //556_Previsión Social
                    celda = celdaMes + "10";
                else if (det.IdConcepto == 557) //557_Impuestos Derivados de la Relación Laboral
                    celda = celdaMes + "11";
                else if (det.IdConcepto == 558) //558_Traslado de Unidades por Terceros
                    celda = celdaMes + "14";
                else if (det.IdConcepto == 559) //559_Gastos por Fletes y Embarques
                    celda = celdaMes + "15";
                else if (det.IdConcepto == 560) //560_Combustibles Y Lubricantes
                    celda = celdaMes + "16";
                else if (det.IdConcepto == 561) //561_Acondicionamiento Unidades Nuevas
                    celda = celdaMes + "17";
                else if (det.IdConcepto == 562) //562_Garantías / Cortesías
                    celda = celdaMes + "18";
                else if (det.IdConcepto == 563) //563_Cuotas y Suscripciones
                    celda = celdaMes + "19";
                else if (det.IdConcepto == 564) //564_Teléfono y Correo
                    celda = celdaMes + "20";
                else if (det.IdConcepto == 565) //565_Publicidad y Promoción
                    celda = celdaMes + "21";
                else if (det.IdConcepto == 566) //566_Herramientas del Taller
                    celda = celdaMes + "22";
                else if (det.IdConcepto == 567) //567_Mantenimiento de Unidades
                    celda = celdaMes + "23";
                else if (det.IdConcepto == 568) //568_Mantenimiento de Equipo
                    celda = celdaMes + "24";
                else if (det.IdConcepto == 569) //569_Materiales Diversos del Taller
                    celda = celdaMes + "25";
                else if (det.IdConcepto == 570) //570_Asesoria Externos
                    celda = celdaMes + "26";
                else if (det.IdConcepto == 571) //571_Capacitación y Adiestramiento
                    celda = celdaMes + "27";
                else if (det.IdConcepto == 572) //572_Papelería y Artículos de Escritorio
                    celda = celdaMes + "28";
                else if (det.IdConcepto == 573) //573_Gastos de Viaje
                    celda = celdaMes + "29";
                else if (det.IdConcepto == 574) //574_Traslado de Valores
                    celda = celdaMes + "30";
                else if (det.IdConcepto == 575) //575_No Deducibles
                    celda = celdaMes + "31";
                else if (det.IdConcepto == 576) //576_Vigilancia y Aseo
                    celda = celdaMes + "32";
                else if (det.IdConcepto == 577) //577_Luz y Agua
                    celda = celdaMes + "33";
                else if (det.IdConcepto == 578) //578_Otros Impuestos y Derechos
                    celda = celdaMes + "34";
                else if (det.IdConcepto == 579) //579_Depreciación y Amortización
                    celda = celdaMes + "37";
                else if (det.IdConcepto == 580) //580_Arrendamiento de Inmuebles
                    celda = celdaMes + "38";
                else if (det.IdConcepto == 581) //581_Seguros y Fianzas
                    celda = celdaMes + "39";
                else if (det.IdConcepto == 582) //582_Publicidad y Promoción Corporativa
                    celda = celdaMes + "40";
                else if (det.IdConcepto == 583) //583_Asesoria y Honorarios
                    celda = celdaMes + "41";
                else if (det.IdConcepto == 584) //584_Arrendamiento de Inmuebles Corporativos
                    celda = celdaMes + "42";
                else if (det.IdConcepto == 585) //585_Mantenimiento de Edificio
                    celda = celdaMes + "43";
                else if (det.IdConcepto == 586) //586_Mejoras al Inmueble
                    celda = celdaMes + "44";

                if (celda.Length >= 2)
                {
                    ws.get_Range(celda, celda).Formula = det.Actual;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + det.Cuenta + " = " + det.Actual);
            }
        }

        public void LlenaPestaniaGRM(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<DetalleSaldo> lstDetSaldo = DetalleSaldo.ListarGeneral(_db, mes, anio, 10, GetAgenciasYSucursales(_db, idAgencia)); //10_REF_REFACCIONES

            celdaMes = GetMes();

            foreach (DetalleSaldo det in lstDetSaldo)
            {
                if (det.IdConcepto == 587) //587_Sueldos y Honorarios Deptos. Operativos
                    celda = celdaMes + "4";
                else if (det.IdConcepto == 588) //588_Comisiones Departamentos Operativos
                    celda = celdaMes + "5";
                else if (det.IdConcepto == 589) //589_Sueldos,  Honorarios, Com. y Prest. Gerentes
                    celda = celdaMes + "6";
                else if (det.IdConcepto == 590) //590_Gratificación Anual Deptos. Operativos
                    celda = celdaMes + "7";
                else if (det.IdConcepto == 591) //591_Sueldos y Honorarios  Depto. Admón.
                    celda = celdaMes + "8";
                else if (det.IdConcepto == 592) //592_Gratificación Anual Depto. Admon
                    celda = celdaMes + "9";
                else if (det.IdConcepto == 593) //593_Previsión Social
                    celda = celdaMes + "10";
                else if (det.IdConcepto == 594) //594_Impuestos Derivados de la Relación Laboral
                    celda = celdaMes + "11";
                else if (det.IdConcepto == 595) //595_Traslado de Unidades por Terceros
                    celda = celdaMes + "14";
                else if (det.IdConcepto == 596) //596_Gastos por Fletes y Embarques
                    celda = celdaMes + "15";
                else if (det.IdConcepto == 597) //597_Combustibles Y Lubricantes
                    celda = celdaMes + "16";
                else if (det.IdConcepto == 598) //598_Acondicionamiento Unidades Nuevas
                    celda = celdaMes + "17";
                else if (det.IdConcepto == 599) //599_Garantías / Cortesías
                    celda = celdaMes + "18";
                else if (det.IdConcepto == 600) //600_Cuotas y Suscripciones
                    celda = celdaMes + "19";
                else if (det.IdConcepto == 601) //601_Teléfono y Correo
                    celda = celdaMes + "20";
                else if (det.IdConcepto == 602) //602_Publicidad y Promoción
                    celda = celdaMes + "21";
                else if (det.IdConcepto == 603) //603_Herramientas del Taller
                    celda = celdaMes + "22";
                else if (det.IdConcepto == 604) //604_Mantenimiento de Unidades
                    celda = celdaMes + "23";
                else if (det.IdConcepto == 605) //605_Mantenimiento de Equipo
                    celda = celdaMes + "24";
                else if (det.IdConcepto == 606) //606_Materiales Diversos del Taller
                    celda = celdaMes + "25";
                else if (det.IdConcepto == 607) //607_Asesoria Externos
                    celda = celdaMes + "26";
                else if (det.IdConcepto == 608) //608_Capacitación y Adiestramiento
                    celda = celdaMes + "27";
                else if (det.IdConcepto == 609) //609_Papelería y Artículos de Escritorio
                    celda = celdaMes + "28";
                else if (det.IdConcepto == 610) //610_Gastos de Viaje
                    celda = celdaMes + "29";
                else if (det.IdConcepto == 611) //611_Traslado de Valores
                    celda = celdaMes + "30";
                else if (det.IdConcepto == 612) //612_No Deducibles
                    celda = celdaMes + "31";
                else if (det.IdConcepto == 613) //613_Vigilancia y Aseo
                    celda = celdaMes + "32";
                else if (det.IdConcepto == 614) //614_Luz y Agua
                    celda = celdaMes + "33";
                else if (det.IdConcepto == 615) //615_Otros Impuestos y Derechos
                    celda = celdaMes + "34";
                else if (det.IdConcepto == 616) //616_Depreciación y Amortización
                    celda = celdaMes + "37";
                else if (det.IdConcepto == 617) //617_Arrendamiento de Inmuebles
                    celda = celdaMes + "38";
                else if (det.IdConcepto == 618) //618_Seguros y Fianzas
                    celda = celdaMes + "39";
                else if (det.IdConcepto == 619) //619_Publicidad y Promoción Corporativa
                    celda = celdaMes + "40";
                else if (det.IdConcepto == 620) //620_Asesoria y Honorarios
                    celda = celdaMes + "41";
                else if (det.IdConcepto == 621) //621_Arrendamiento de Inmuebles Corporativos
                    celda = celdaMes + "42";
                else if (det.IdConcepto == 622) //622_Mantenimiento de Edificio
                    celda = celdaMes + "43";
                else if (det.IdConcepto == 623) //623_Mejoras al Inmueble
                    celda = celdaMes + "44";

                if (celda.Length >= 2)
                {
                    ws.get_Range(celda, celda).Formula = det.Actual;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + det.Cuenta + " = " + det.Actual);
            }
        }

        public void LlenaPestaniaGDM(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            List<DetalleSaldo> lstDetSaldo = DetalleSaldo.ListarGeneral(_db, mes, anio, 9, GetAgenciasYSucursales(_db, idAgencia)); //9_ADM_ADMINISTRACION

            celdaMes = GetMes();

            PeriodoContable periodo = PeriodoContable.BuscarPorMesAnio(_db, mes, anio);

            SaldoPorPeriodoPorCuentaBalanza saldo6000001000070495 = SaldoPorPeriodoPorCuentaBalanza.Buscar(_db, idAgencia, periodo.Id, 1, 166673); //166673_6000001000070495_GASTOS || ADMINISTRACION || GASTOS GENERALES || INGRESOS ASIMILADOS A SALARIOS DG ||

            decimal saldoCuenta = 0;

            if (saldo6000001000070495 != null)
                saldoCuenta += Convert.ToInt32(saldo6000001000070495.TotalDeCargos - saldo6000001000070495.TotalDeAbonos);

            foreach (DetalleSaldo det in lstDetSaldo)
            {
                if (det.IdConcepto == 624) //624_Sueldos y Honorarios Deptos. Operativos
                    celda = celdaMes + "4";
                else if (det.IdConcepto == 625) //625_Comisiones Departamentos Operativos
                    celda = celdaMes + "5";
                else if (det.IdConcepto == 626) //626_Sueldos,  Honorarios, Com. y Prest. Gerentes
                {
                    celda = celdaMes + "6";

                    if (version == "V2")
                        det.Actual = det.Actual - Math.Round(saldoCuenta / 1000, 0);
                }
                else if (det.IdConcepto == 627) //627_Gratificación Anual Deptos. Operativos
                    celda = celdaMes + "7";
                else if (det.IdConcepto == 628) //628_Sueldos y Honorarios  Depto. Admón.
                    celda = celdaMes + "8";
                else if (det.IdConcepto == 629) //629_Gratificación Anual Depto. Admon
                    celda = celdaMes + "9";
                else if (det.IdConcepto == 630) //630_Previsión Social
                    celda = celdaMes + "10";
                else if (det.IdConcepto == 631) //631_Impuestos Derivados de la Relación Laboral
                    celda = celdaMes + "11";
                else if (det.IdConcepto == 632) //632_Traslado de Unidades por Terceros
                    celda = celdaMes + "14";
                else if (det.IdConcepto == 633) //633_Gastos por Fletes y Embarques
                    celda = celdaMes + "15";
                else if (det.IdConcepto == 634) //634_Combustibles Y Lubricantes
                    celda = celdaMes + "16";
                else if (det.IdConcepto == 635) //635_Acondicionamiento Unidades Nuevas
                    celda = celdaMes + "17";
                else if (det.IdConcepto == 636) //636_Garantías / Cortesías
                    celda = celdaMes + "18";
                else if (det.IdConcepto == 637) //637_Cuotas y Suscripciones
                    celda = celdaMes + "19";
                else if (det.IdConcepto == 638) //638_Teléfono y Correo
                    celda = celdaMes + "20";
                else if (det.IdConcepto == 639) //639_Publicidad y Promoción
                    celda = celdaMes + "21";
                else if (det.IdConcepto == 640) //640_Herramientas del Taller
                    celda = celdaMes + "22";
                else if (det.IdConcepto == 641) //641_Mantenimiento de Unidades
                    celda = celdaMes + "23";
                else if (det.IdConcepto == 642) //642_Mantenimiento de Equipo
                    celda = celdaMes + "24";
                else if (det.IdConcepto == 643) //643_Materiales Diversos del Taller
                    celda = celdaMes + "25";
                else if (det.IdConcepto == 644) //644_Asesoria Externos
                    celda = celdaMes + "26";
                else if (det.IdConcepto == 645) //645_Capacitación y Adiestramiento
                    celda = celdaMes + "27";
                else if (det.IdConcepto == 646) //646_Papelería y Artículos de Escritorio
                    celda = celdaMes + "28";
                else if (det.IdConcepto == 647) //647_Gastos de Viaje
                    celda = celdaMes + "29";
                else if (det.IdConcepto == 648) //648_Traslado de Valores
                    celda = celdaMes + "30";
                else if (det.IdConcepto == 649) //649_No Deducibles
                    celda = celdaMes + "31";
                else if (det.IdConcepto == 650) //650_Vigilancia y Aseo
                    celda = celdaMes + "32";
                else if (det.IdConcepto == 651) //651_Luz y Agua
                    celda = celdaMes + "33";
                else if (det.IdConcepto == 652) //652_Otros Impuestos y Derechos
                    celda = celdaMes + "34";
                else if (det.IdConcepto == 653) //653_Depreciación y Amortización
                    celda = celdaMes + "37";
                else if (det.IdConcepto == 654) //654_Arrendamiento de Inmuebles
                    celda = celdaMes + "38";
                else if (det.IdConcepto == 655) //655_Seguros y Fianzas
                    celda = celdaMes + "39";
                else if (det.IdConcepto == 656) //656_Publicidad y Promoción Corporativa
                    celda = celdaMes + "40";
                else if (det.IdConcepto == 657) //657_Asesoria y Honorarios
                    celda = celdaMes + "41";
                else if (det.IdConcepto == 658) //658_Arrendamiento de Inmuebles Corporativos
                    celda = celdaMes + "42";
                else if (det.IdConcepto == 659) //659_Mantenimiento de Edificio
                    celda = celdaMes + "43";
                else if (det.IdConcepto == 660) //660_Mejoras al Inmueble
                    celda = celdaMes + "44";

                if (celda.Length >= 2)
                {
                    ws.get_Range(celda, celda).Formula = det.Actual;
                }

                celda = "";

                Console.WriteLine("[CONCEPTO]: " + det.Cuenta + " = " + det.Actual);
            }
        }

        public void AjustaPestaniasGastos(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            celdaMes = GetMes();

            celda = celdaMes + "49";

            string cadena = "";
            decimal diferencia = 0;

            //cadena = ws.Cells[49, 5];
            //cadena = Convert.ToString(ws.get_Range(celda).Value);

            if (Convert.ToString(ws.get_Range(celda).Value) != "")
            {
                cadena = Convert.ToString(ws.get_Range(celda).Value);
                diferencia = Convert.ToDecimal(ws.get_Range(celda).Value);
            }

            Console.WriteLine("[CONCEPTO]: diferencia = " + diferencia);

            celda = celdaMes + "31";

            decimal valor = 0;

            if (Convert.ToString(ws.get_Range(celda).Value) != "")
            {
                cadena = Convert.ToString(ws.get_Range(celda).Value);
                valor = Convert.ToDecimal(ws.get_Range(celda).Value);
            }

            Console.WriteLine("[CONCEPTO]: valor = " + valor);

            decimal ajuste = 0;

            if ((diferencia > 0 && diferencia <= 5) || (diferencia < 0 && diferencia >= -5))
            {
                celda = celdaMes + "31";

                ajuste = valor + (diferencia * -1);

                ws.get_Range(celda, celda).Formula = ajuste;
            }

            Console.WriteLine("[CONCEPTO]: ajuste = " + ajuste);
        }

        public void AjustaPestaniaRM(ExcelApp.Worksheet ws)
        {
            string celda = "";
            string celdaMes = "";

            celdaMes = GetMes();

            celda = celdaMes + "77";

            string strGastosAdmin = ""; //1044_GASTOS ADMINISTRATIVOS
            decimal decGastosAdmin = 0;

            if (Convert.ToString(ws.get_Range(celda).Value) != "")
            {
                strGastosAdmin = Convert.ToString(ws.get_Range(celda).Value);
                decGastosAdmin = Convert.ToDecimal(ws.get_Range(celda).Value);
            }

            Console.WriteLine("[CONCEPTO]: 1044_GASTOS ADMINISTRATIVOS = " + decGastosAdmin);

            celda = celdaMes + "13";

            string strGastosAdminAN = ""; //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS
            decimal decGastosAdminAN = 0;

            if (Convert.ToString(ws.get_Range(celda).Value) != "")
            {
                strGastosAdminAN = Convert.ToString(ws.get_Range(celda).Value);
                decGastosAdminAN = Convert.ToDecimal(ws.get_Range(celda).Value);
            }

            Console.WriteLine("[CONCEPTO]: 1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS = " + decGastosAdminAN);

            celda = celdaMes + "19";

            string strGastosAdminGN = ""; //1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS
            decimal decGastosAdminGN = 0;

            if (Convert.ToString(ws.get_Range(celda).Value) != "")
            {
                strGastosAdminGN = Convert.ToString(ws.get_Range(celda).Value);
                decGastosAdminGN = Convert.ToDecimal(ws.get_Range(celda).Value);
            }

            Console.WriteLine("[CONCEPTO]: 1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS = " + decGastosAdminGN);

            celda = celdaMes + "29";

            string strGastosAdminAU = ""; //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS
            decimal decGastosAdminAU = 0;

            if (Convert.ToString(ws.get_Range(celda).Value) != "")
            {
                strGastosAdminAU = Convert.ToString(ws.get_Range(celda).Value);
                decGastosAdminAU = Convert.ToDecimal(ws.get_Range(celda).Value);
            }

            Console.WriteLine("[CONCEPTO]: 1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS = " + decGastosAdminAU);

            celda = celdaMes + "35";

            string strGastosAdminSE = ""; //1091_GASTOS DE ADMINISTRACION SERVICIO
            decimal decGastosAdminSE = 0;

            if (Convert.ToString(ws.get_Range(celda).Value) != "")
            {
                strGastosAdminSE = Convert.ToString(ws.get_Range(celda).Value);
                decGastosAdminSE = Convert.ToDecimal(ws.get_Range(celda).Value);
            }

            Console.WriteLine("[CONCEPTO]: 1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS = " + decGastosAdminSE);

            celda = celdaMes + "41";

            string strGastosAdminHyP = ""; //1092_GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA
            decimal decGastosAdminHyP = 0;

            if (Convert.ToString(ws.get_Range(celda).Value) != "")
            {
                strGastosAdminHyP = Convert.ToString(ws.get_Range(celda).Value);
                decGastosAdminHyP = Convert.ToDecimal(ws.get_Range(celda).Value);
            }

            Console.WriteLine("[CONCEPTO]: 1092_GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA = " + decGastosAdminHyP);

            celda = celdaMes + "58";

            string strGastosAdminRE = ""; //1093_GASTOS DE ADMINISTRACION REFACCIONES
            decimal decGastosAdminRE = 0;

            if (Convert.ToString(ws.get_Range(celda).Value) != "")
            {
                strGastosAdminRE = Convert.ToString(ws.get_Range(celda).Value);
                decGastosAdminRE = Convert.ToDecimal(ws.get_Range(celda).Value);
            }

            Console.WriteLine("[CONCEPTO]: 1093_GASTOS DE ADMINISTRACION REFACCIONES = " + decGastosAdminRE);

            decimal gastosAdminProrrateados = decGastosAdminAN + decGastosAdminGN + decGastosAdminAU + decGastosAdminSE + decGastosAdminHyP + decGastosAdminRE;
            decimal diferencia = decGastosAdmin - gastosAdminProrrateados;
            decimal ajuste = 0;

            Console.WriteLine("[CONCEPTO]: gastosAdminProrrateados = " + gastosAdminProrrateados);
            Console.WriteLine("[CONCEPTO]: diferencia = " + diferencia);

            if ((diferencia > 0 && diferencia <= 6))
            {
                if (diferencia == 1)
                {
                    celda = celdaMes + "13"; //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS

                    ajuste = decGastosAdminAN + diferencia;

                    ws.get_Range(celda, celda).Formula = ajuste;
                }
                else if (diferencia == 2)
                {
                    celda = celdaMes + "13"; //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS

                    ajuste = decGastosAdminAN + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "29"; //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS

                    ajuste = decGastosAdminAU + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;
                }
                else if (diferencia == 3)
                {
                    celda = celdaMes + "13"; //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS

                    ajuste = decGastosAdminAN + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "29"; //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS

                    ajuste = decGastosAdminAU + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "19"; //1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS

                    ajuste = decGastosAdminGN + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;
                }
                else if (diferencia == 4)
                {
                    celda = celdaMes + "13"; //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS

                    ajuste = decGastosAdminAN + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "29"; //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS

                    ajuste = decGastosAdminAU + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "19"; //1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS

                    ajuste = decGastosAdminGN + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "58"; //1093_GASTOS DE ADMINISTRACION REFACCIONES

                    ajuste = decGastosAdminRE + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;
                }
                else if (diferencia == 5)
                {
                    celda = celdaMes + "13"; //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS

                    ajuste = decGastosAdminAN + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "29"; //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS

                    ajuste = decGastosAdminAU + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "19"; //1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS

                    ajuste = decGastosAdminGN + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "58"; //1093_GASTOS DE ADMINISTRACION REFACCIONES

                    ajuste = decGastosAdminRE + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "35"; //1091_GASTOS DE ADMINISTRACION SERVICIO

                    ajuste = decGastosAdminSE + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;
                }
                else if (diferencia == 6)
                {
                    celda = celdaMes + "13"; //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS

                    ajuste = decGastosAdminAN + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "29"; //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS

                    ajuste = decGastosAdminAU + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "19"; //1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS

                    ajuste = decGastosAdminGN + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "58"; //1093_GASTOS DE ADMINISTRACION REFACCIONES

                    ajuste = decGastosAdminRE + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "35"; //1091_GASTOS DE ADMINISTRACION SERVICIO

                    ajuste = decGastosAdminSE + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;

                    celda = celdaMes + "41"; //1092_GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA

                    ajuste = decGastosAdminHyP + 1;

                    ws.get_Range(celda, celda).Formula = ajuste;
                }
            }
        }

        public string GetMes()
        {
            string sMes = "";

            switch (pestania)
            {
                case "RM":
                case "RMAN":
                case "BG":
                case "AG":
                case "AF":
                case "VP":
                case "GANMT":
                case "GANMF":
                case "GANMVI":
                case "GGNM":
                case "GASM":
                case "GSM":
                case "GHPM":
                case "GRM":
                case "GDM":
                    {
                        for (int i = 1; i <= 12; i++)
                        {
                            if (i == 1 && i == mes)
                                sMes = "B";
                            else if (i == 2 && i == mes)
                                sMes = "C";
                            else if (i == 3 && i == mes)
                                sMes = "D";
                            else if (i == 4 && i == mes)
                                sMes = "E";
                            else if (i == 5 && i == mes)
                                sMes = "F";
                            else if (i == 6 && i == mes)
                                sMes = "G";
                            else if (i == 7 && i == mes)
                                sMes = "H";
                            else if (i == 8 && i == mes)
                                sMes = "I";
                            else if (i == 9 && i == mes)
                                sMes = "J";
                            else if (i == 10 && i == mes)
                                sMes = "K";
                            else if (i == 11 && i == mes)
                                sMes = "L";
                            else if (i == 12 && i == mes)
                                sMes = "M";
                        }
                        break;
                    }
                case "PX":
                case "Cuentas":                
                    {
                        for (int i = 1; i <= 12; i++)
                        {
                            if (i == 1 && i == mes)
                                sMes = "C";
                            else if (i == 2 && i == mes)
                                sMes = "D";
                            else if (i == 3 && i == mes)
                                sMes = "E";
                            else if (i == 4 && i == mes)
                                sMes = "F";
                            else if (i == 5 && i == mes)
                                sMes = "G";
                            else if (i == 6 && i == mes)
                                sMes = "H";
                            else if (i == 7 && i == mes)
                                sMes = "I";
                            else if (i == 8 && i == mes)
                                sMes = "J";
                            else if (i == 9 && i == mes)
                                sMes = "K";
                            else if (i == 10 && i == mes)
                                sMes = "L";
                            else if (i == 11 && i == mes)
                                sMes = "M";
                            else if (i == 12 && i == mes)
                                sMes = "N";
                        }
                        break;
                    }
                //case "VP":
                //    {
                //        for (int i = 1; i <= 12; i++)
                //        {
                //            if (i == 1 && i == mes)
                //                sMes = "B";
                //            else if (i == 2 && i == mes)
                //                sMes = "C";
                //            else if (i == 3 && i == mes)
                //                sMes = "D";
                //            else if (i == 4 && i == mes)
                //                sMes = "E";
                //            else if (i == 5 && i == mes)
                //                sMes = "F";
                //            else if (i == 6 && i == mes)
                //                sMes = "G";
                //            else if (i == 7 && i == mes)
                //                sMes = "H";
                //            else if (i == 8 && i == mes)
                //                sMes = "I";
                //            else if (i == 9 && i == mes)
                //                sMes = "J";
                //            else if (i == 10 && i == mes)
                //                sMes = "K";
                //            else if (i == 11 && i == mes)
                //                sMes = "L";
                //            else if (i == 12 && i == mes)
                //                sMes = "M";
                //        }
                //        break;
                //    }
                //case "RMAN":
                //    {
                //        for (int i = 1; i <= 12; i++)
                //        {
                //            if (i == 1 && i == mes)
                //                sMes = "B";
                //            else if (i == 2 && i == mes)
                //                sMes = "C";
                //            else if (i == 3 && i == mes)
                //                sMes = "D";
                //            else if (i == 4 && i == mes)
                //                sMes = "E";
                //            else if (i == 5 && i == mes)
                //                sMes = "F";
                //            else if (i == 6 && i == mes)
                //                sMes = "G";
                //            else if (i == 7 && i == mes)
                //                sMes = "H";
                //            else if (i == 8 && i == mes)
                //                sMes = "I";
                //            else if (i == 9 && i == mes)
                //                sMes = "J";
                //            else if (i == 10 && i == mes)
                //                sMes = "K";
                //            else if (i == 11 && i == mes)
                //                sMes = "L";
                //            else if (i == 12 && i == mes)
                //                sMes = "M";
                //        }
                //        break;
                //    }
                //case "RM":
                //    {
                //        for (int i = 1; i <= 12; i++)
                //        {
                //            if (i == 1 && i == mes)
                //                sMes = "B";
                //            else if (i == 2 && i == mes)
                //                sMes = "C";
                //            else if (i == 3 && i == mes)
                //                sMes = "D";
                //            else if (i == 4 && i == mes)
                //                sMes = "E";
                //            else if (i == 5 && i == mes)
                //                sMes = "F";
                //            else if (i == 6 && i == mes)
                //                sMes = "G";
                //            else if (i == 7 && i == mes)
                //                sMes = "H";
                //            else if (i == 8 && i == mes)
                //                sMes = "I";
                //            else if (i == 9 && i == mes)
                //                sMes = "J";
                //            else if (i == 10 && i == mes)
                //                sMes = "K";
                //            else if (i == 11 && i == mes)
                //                sMes = "L";
                //            else if (i == 12 && i == mes)
                //                sMes = "M";
                //        }
                //        break;
                //    }
                case "OPL":
                    {
                        for (int i = 1; i <= 12; i++)
                        {
                            if (i == 1 && i == mes)
                                sMes = "D";
                            else if (i == 2 && i == mes)
                                sMes = "E";
                            else if (i == 3 && i == mes)
                                sMes = "F";
                            else if (i == 4 && i == mes)
                                sMes = "G";
                            else if (i == 5 && i == mes)
                                sMes = "H";
                            else if (i == 6 && i == mes)
                                sMes = "I";
                            else if (i == 7 && i == mes)
                                sMes = "J";
                            else if (i == 8 && i == mes)
                                sMes = "K";
                            else if (i == 9 && i == mes)
                                sMes = "L";
                            else if (i == 10 && i == mes)
                                sMes = "M";
                            else if (i == 11 && i == mes)
                                sMes = "N";
                            else if (i == 12 && i == mes)
                                sMes = "O";
                        }
                        break;
                    }
                //case "PX":
                //    {
                //        for (int i = 1; i <= 12; i++)
                //        {
                //            if (i == 1 && i == mes)
                //                sMes = "C";
                //            else if (i == 2 && i == mes)
                //                sMes = "D";
                //            else if (i == 3 && i == mes)
                //                sMes = "E";
                //            else if (i == 4 && i == mes)
                //                sMes = "F";
                //            else if (i == 5 && i == mes)
                //                sMes = "G";
                //            else if (i == 6 && i == mes)
                //                sMes = "H";
                //            else if (i == 7 && i == mes)
                //                sMes = "I";
                //            else if (i == 8 && i == mes)
                //                sMes = "J";
                //            else if (i == 9 && i == mes)
                //                sMes = "K";
                //            else if (i == 10 && i == mes)
                //                sMes = "L";
                //            else if (i == 11 && i == mes)
                //                sMes = "M";
                //            else if (i == 12 && i == mes)
                //                sMes = "N";
                //        }
                //        break;
                //    }
                case "SI":
                    {
                        for (int i = 1; i <= 12; i++)
                        {
                            if (i == 1 && i == mes)
                                sMes = "6";
                            else if (i == 2 && i == mes)
                                sMes = "21";
                            else if (i == 3 && i == mes)
                                sMes = "36";
                            else if (i == 4 && i == mes)
                                sMes = "51";
                            else if (i == 5 && i == mes)
                                sMes = "66";
                            else if (i == 6 && i == mes)
                                sMes = "81";
                            else if (i == 7 && i == mes)
                                sMes = "96";
                            else if (i == 8 && i == mes)
                                sMes = "111";
                            else if (i == 9 && i == mes)
                                sMes = "126";
                            else if (i == 10 && i == mes)
                                sMes = "141";
                            else if (i == 11 && i == mes)
                                sMes = "156";
                            else if (i == 12 && i == mes)
                                sMes = "171";
                        }
                        break;
                    }
                case "SC":
                    {
                        for (int i = 1; i <= 12; i++)
                        {
                            if (i == 1 && i == mes)
                                sMes = "8";
                            else if (i == 2 && i == mes)
                                sMes = "27";
                            else if (i == 3 && i == mes)
                                sMes = "46";
                            else if (i == 4 && i == mes)
                                sMes = "65";
                            else if (i == 5 && i == mes)
                                sMes = "84";
                            else if (i == 6 && i == mes)
                                sMes = "103";
                            else if (i == 7 && i == mes)
                                sMes = "122";
                            else if (i == 8 && i == mes)
                                sMes = "141";
                            else if (i == 9 && i == mes)
                                sMes = "160";
                            else if (i == 10 && i == mes)
                                sMes = "179";
                            else if (i == 11 && i == mes)
                                sMes = "198";
                            else if (i == 12 && i == mes)
                                sMes = "217";
                        }
                        break;
                    }
                case "CC":
                    {
                        for (int i = 1; i <= 12; i++)
                        {
                            if (i == 1 && i == mes)
                                sMes = "7";
                            else if (i == 2 && i == mes)
                                sMes = "8";
                            else if (i == 3 && i == mes)
                                sMes = "9";
                            else if (i == 4 && i == mes)
                                sMes = "10";
                            else if (i == 5 && i == mes)
                                sMes = "11";
                            else if (i == 6 && i == mes)
                                sMes = "12";
                            else if (i == 7 && i == mes)
                                sMes = "13";
                            else if (i == 8 && i == mes)
                                sMes = "14";
                            else if (i == 9 && i == mes)
                                sMes = "15";
                            else if (i == 10 && i == mes)
                                sMes = "16";
                            else if (i == 11 && i == mes)
                                sMes = "17";
                            else if (i == 12 && i == mes)
                                sMes = "18";
                        }
                        break;
                    }
                //case "GANMT":
                //case "GGNM":
                //case "GASM":
                //case "GSM":
                //case "GHPM":
                //case "GRM":
                //case "GDM":
                //    {
                //        for (int i = 1; i <= 12; i++)
                //        {
                //            if (i == 1 && i == mes)
                //                sMes = "B";
                //            else if (i == 2 && i == mes)
                //                sMes = "C";
                //            else if (i == 3 && i == mes)
                //                sMes = "D";
                //            else if (i == 4 && i == mes)
                //                sMes = "E";
                //            else if (i == 5 && i == mes)
                //                sMes = "F";
                //            else if (i == 6 && i == mes)
                //                sMes = "G";
                //            else if (i == 7 && i == mes)
                //                sMes = "H";
                //            else if (i == 8 && i == mes)
                //                sMes = "I";
                //            else if (i == 9 && i == mes)
                //                sMes = "J";
                //            else if (i == 10 && i == mes)
                //                sMes = "K";
                //            else if (i == 11 && i == mes)
                //                sMes = "L";
                //            else if (i == 12 && i == mes)
                //                sMes = "M";
                //        }
                //        break;
                //    }
            }
                
            return sMes;
        }

        private static List<int> GetAgenciasYSucursales(DB2Database aDB, int aIdAgencia)
        {
            List<int> lstIdsAgencias = new List<int>();

            System.Data.DataTable dt = aDB.GetDataTable(@"SELECT SUC.FIFNIDCIAS AS SUCURSALES
                        FROM [PREFIX]FINA.FNDAGSUC SUC
                        INNER JOIN [PREFIX]GRAL.GECCIAUN AG ON AG.FIGEIDCIAU = SUC.FIFNIDCIAU AND AG.FIGESTATUS = 1
                        INNER JOIN [PREFIX]GRAL.GECCIAUN AGSUC ON AGSUC.FIGEIDCIAU = SUC.FIFNIDCIAS AND AGSUC.FIGESTATUS = 1 
                        WHERE FIFNSTATUS = 1 and SUC.FIFNIDCIAU = " + aIdAgencia);

            lstIdsAgencias.Add(aIdAgencia);

            if (dt.Rows.Count > 0)
                if (dt.Rows[0]["SUCURSALES"].ToString().Trim() != "")
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        lstIdsAgencias.Add(Convert.ToInt32(dr["SUCURSALES"]));
                    }

                    return lstIdsAgencias;
                }
                else
                    return lstIdsAgencias;
            else
                return lstIdsAgencias;
        }
    }
}
