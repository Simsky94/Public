using DVAModelsReflection;
using DVAModelsReflectionFINA.Models.FINA;
using DVAModelsReflectionFINA1.Models.FINA;

//using DVAModelsReflectionFINA1.Models.FINA;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
//using DVAExcel;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    public class ProcesoDeLecturaArchivoBaseExcelYLlenadoInfo
    {
        DB2Database _db = null;
        

        int idAgencia = 0;
        int anio = 0;
        int mes = 0;
        string siglas = "";
        string version = "";
        string pestania = "";
        List<ConceptosContables> liConceptos = new List<ConceptosContables>();
        string ruta = @"D:\Users\jasoria\Desktop\Escritorio\PMO\Proyectos 2023\ARCHIVO BASE\ARCHIVOS BASE EXCEL\";
        //string ruta = @"E:\ARCHIVOS BASE EXCEL\";
        //string ruta = @"C:\Users\fevangelista\Documents\Reportes Cuaderno Contable\";

        public ProcesoDeLecturaArchivoBaseExcelYLlenadoInfo(DB2Database _db, int aAnio, int aMes, int aIdAgencia, string aVersion, string aSiglas, string aPestania)
        {
            this._db = _db;
            this.idAgencia = aIdAgencia;
            this.mes = aMes;
            this.anio = aAnio;
            this.version = aVersion;
            this.siglas = aSiglas;
            this.pestania = aPestania;
            ruta = ruta + anio + "\\" + mes + "\\" + version + "\\ARCHIVO_BASE_EXCEL\\";
            liConceptos = ConceptosContables.ListarRM(_db);
            Console.WriteLine("[INICIA LECTURA ARCHIVO BASE EXCEL PESTAÑA " + pestania + "]: ");
            Console.WriteLine("[ID_AGENCIA]: " + idAgencia);
            Console.WriteLine("[AÑO]: " + anio);
            Console.WriteLine("[MES]: " + mes);
            Console.WriteLine("[VERSION]: " + version);
            Console.WriteLine("[SIGLAS]: " + siglas);
            Console.WriteLine("[PESTAÑA]: " + pestania);
        }

        public void LecturaExcel()
        {   
            try
            {
                if (version == "V1")
                    ruta = ruta + "ARCHIVO BASE " + anio + " " + siglas + ".xls";
                else
                    ruta = ruta + "ARCHIVO BASE " + anio + " VERSION 2 " + siglas + ".xls";

                DVAExcel.ExcelReader eR = new DVAExcel.ExcelReader(ruta);

                //List<DVAExcel.HojaDeTrabajo> hojas = eR.GetHojas();

                DVAExcel.HojaDeTrabajo hojaDeTrabajo = eR.GetHoja(pestania);

                _db = new DB2Database();
                _db.BeginTransaction();

                //foreach (DVAExcel.HojaDeTrabajo hojaDeTrabajo in hojas)
                //{
                if (hojaDeTrabajo.Nombre == "RM")
                {
                    LeePestanaRM(hojaDeTrabajo);
                }
                else if (hojaDeTrabajo.Nombre == "BG")
                {
                    LeePestanaBG(hojaDeTrabajo);
                }
                else if (hojaDeTrabajo.Nombre == "VP")
                {
                    LeerPestanaVP(hojaDeTrabajo);
                }
                else if (hojaDeTrabajo.Nombre == "Cuentas")
                {
                    LeePestanaCuentas(hojaDeTrabajo);
                }
                else if (hojaDeTrabajo.Nombre == "AG")
                {
                    LeePestanaAG(hojaDeTrabajo);
                }
                else if (hojaDeTrabajo.Nombre == "SC")
                {
                    LeePestanaSC(hojaDeTrabajo);
                }
                //}                

                _db.CommitTransaction();

                eR.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine (ex.ToString());

                _db.RollbackTransaction();
            }
        }

        public void LeePestanaRM(DVAExcel.HojaDeTrabajo hoja)
        {
            string valor = "";
            int r = 1;
            int c = 1;

            ProcesoResultadoMensualExcel v1;
            ProcesoResultadoMensualExtralibrosExcel v2;

            foreach (object celda in hoja.Celdas) 
            {
                if (celda != null)
                    valor = celda.ToString();

                if ((c == (mes + 1)) && r >= 4)
                {
                    if (r == 4) //1000_INGRESO TOTAL
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1000;
                            v1.Concepto = "INGRESO TOTAL";                            
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1000;
                            v2.Concepto = "INGRESO TOTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 5) //1001_UNIDADES NUEVAS VENDIDAS
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1001;
                            v1.Concepto = "UNIDADES NUEVAS VENDIDAS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1001;
                            v2.Concepto = "UNIDADES NUEVAS VENDIDAS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 6) //1002_INGRESOS POR VENTA DE UNIDADES NUEVAS
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1002;
                            v1.Concepto = "INGRESOS POR VENTA DE UNIDADES NUEVAS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1002;
                            v2.Concepto = "INGRESOS POR VENTA DE UNIDADES NUEVAS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 7) //1003_PROMEDIO DE UTILIDAD BRUTA POR UNIDAD
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1003;
                            v1.Concepto = "PROMEDIO DE UTILIDAD BRUTA POR UNIDAD";                            
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1003;
                            v2.Concepto = "PROMEDIO DE UTILIDAD BRUTA POR UNIDAD";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 8) //1004_UTILIDAD BRUTA DEPARTAMENTA
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1004;
                            v1.Concepto = "UTILIDAD BRUTA DEPARTAMENTA";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1004;
                            v2.Concepto = "UTILIDAD BRUTA DEPARTAMENTA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 9) //1005_GASTOS DEPARTAMENTALES
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1005;
                            v1.Concepto = "GASTOS DEPARTAMENTALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1005;
                            v2.Concepto = "GASTOS DEPARTAMENTALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 10) 
                    {
                        if (version == "V1") //1006_UTILIDAD NETA AUTOS NUEVOS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1006;
                            v1.Concepto = "UTILIDAD NETA AUTOS NUEVOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1007_UTILIDAD NETA SERVICIOS ADICIONALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1007;
                            v2.Concepto = "UTILIDAD NETA SERVICIOS ADICIONALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 11)
                    {
                        if (version == "V1") //1007_UTILIDAD NETA SERVICIOS ADICIONALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1007;
                            v1.Concepto = "UTILIDAD NETA SERVICIOS ADICIONALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1008_UTILIDAD NETA DEPARTAMENTAL
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1008;
                            v2.Concepto = "UTILIDAD NETA DEPARTAMENTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 12)
                    {
                        if (version == "V1") //1008_UTILIDAD NETA DEPARTAMENTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1008;
                            v1.Concepto = "UTILIDAD NETA DEPARTAMENTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1085_FINANCIAMIENTO NETO AUTOS NUEVOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1085;
                            v2.Concepto = "FINANCIAMIENTO NETO AUTOS NUEVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 13)
                    {
                        if (version == "V1") //1009_INGRESOS GERENCIA DE NEGOCIOS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1009;
                            v1.Concepto = "INGRESOS GERENCIA DE NEGOCIOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1088_GASTOS DE ADMINISTRACION AUTOS NUEVOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1088;
                            v2.Concepto = "GASTOS DE ADMINISTRACION AUTOS NUEVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 14)
                    {
                        if (version == "V1") //1010_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1010;
                            v1.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //2024_UTILIDAD NETA DEPARTAMENTAL AUTOS NUEVOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 2024;
                            v2.Concepto = "UTILIDAD NETA DEPARTAMENTAL AUTOS NUEVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 15)
                    {
                        if (version == "V1") //1011_GASTOS DEPARTAMENTALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1011;
                            v1.Concepto = "GASTOS DEPARTAMENTALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1009_INGRESOS GERENCIA DE NEGOCIOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1009;
                            v2.Concepto = "INGRESOS GERENCIA DE NEGOCIOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 16)
                    {
                        if (version == "V1") //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1012;
                            v1.Concepto = "UTILIDAD NETA GERENCIA DE NEGOCIOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1010_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1010;
                            v2.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 17)
                    {
                        if (version == "V1") //1013_UNIDADES SEMINUEVAS VENDIDAS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1013;
                            v1.Concepto = "UNIDADES SEMINUEVAS VENDIDAS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1011_GASTOS DEPARTAMENTALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1011;
                            v2.Concepto = "GASTOS DEPARTAMENTALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 18)
                    {
                        if (version == "V1") //1014_INGRESOS VENTA DE UNIDADES SEMINUEVAS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1014;
                            v1.Concepto = "INGRESOS VENTA DE UNIDADES SEMINUEVAS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1012_UTILIDAD NETA GERENCIA DE NEGOCIOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1012;
                            v2.Concepto = "UTILIDAD NETA GERENCIA DE NEGOCIOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 19)
                    {
                        if (version == "V1") //1015_PROMEDIO DE UTILIDAD BRUTA POR UNIDAD
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1015;
                            v1.Concepto = "PROMEDIO DE UTILIDAD BRUTA POR UNIDAD";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0)), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1089_GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1089;
                            v2.Concepto = "GASTOS DE ADMINISTRACION GERENCIA DE NEGOCIOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 20)
                    {
                        if (version == "V1") //1016_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1016;
                            v1.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //2025_UTILIDAD NETA DEPARTAMENTAL GERENCIA DE NEGOCIOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 2025;
                            v2.Concepto = "UTILIDAD NETA DEPARTAMENTAL GERENCIA DE NEGOCIOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 21)
                    {
                        if (version == "V1") //1017_GASTOS DEPARTAMENTALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1017;
                            v1.Concepto = "GASTOS DEPARTAMENTALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1013_UNIDADES SEMINUEVAS VENDIDAS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1013;
                            v2.Concepto = "UNIDADES SEMINUEVAS VENDIDAS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 22)
                    {
                        if (version == "V1") //1018_UTILIDAD NETA AUTOS SEMINUEVOS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1018;
                            v1.Concepto = "UTILIDAD NETA AUTOS SEMINUEVOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1014_INGRESOS VENTA DE UNIDADES SEMINUEVAS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1014;
                            v2.Concepto = "INGRESOS VENTA DE UNIDADES SEMINUEVAS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 23)
                    {
                        if (version == "V1") //1019_UTILIDAD NETA SERVICIOS ADICIONALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1019;
                            v1.Concepto = "UTILIDAD NETA SERVICIOS ADICIONALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1015_PROMEDIO DE UTILIDAD BRUTA POR UNIDAD
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1015;
                            v2.Concepto = "PROMEDIO DE UTILIDAD BRUTA POR UNIDAD";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0)), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 24)
                    {
                        if (version == "V1") //1020_UTILIDAD NETA DEPARTAMENTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1020;
                            v1.Concepto = "UTILIDAD NETA DEPARTAMENTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1016_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1016;
                            v2.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 25)
                    {
                        if (version == "V1") //1021_INGRESOS POR VENTA DE SERVICIO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1021;
                            v1.Concepto = "INGRESOS POR VENTA DE SERVICIO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1017_GASTOS DEPARTAMENTALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1017;
                            v2.Concepto = "GASTOS DEPARTAMENTALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 26)
                    {
                        if (version == "V1") //1022_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1022;
                            v1.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1019_UTILIDAD NETA SERVICIOS ADICIONALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1019;
                            v2.Concepto = "UTILIDAD NETA SERVICIOS ADICIONALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 27)
                    {
                        if (version == "V1") //1023_GASTOS DEPARTAMENTALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1023;
                            v1.Concepto = "GASTOS DEPARTAMENTALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1020_UTILIDAD NETA DEPARTAMENTAL
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1020;
                            v2.Concepto = "UTILIDAD NETA DEPARTAMENTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 28)
                    {
                        if (version == "V1") //1024_UTILIDAD NETA SERVICIO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1024;
                            v1.Concepto = "UTILIDAD NETA SERVICIO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1086_FINANCIAMIENTO NETO AUTOS SEMINUEVOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1086;
                            v2.Concepto = "FINANCIAMIENTO NETO AUTOS SEMINUEVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 29)
                    {
                        if (version == "V1") //1025_INGRESOS POR HOJALATERIA Y PINTURA
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1025;
                            v1.Concepto = "INGRESOS POR HOJALATERIA Y PINTURA";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1090_GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1090;
                            v2.Concepto = "GASTOS DE ADMINISTRACION AUTOS SEMINUEVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 30)
                    {
                        if (version == "V1") //1026_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1026;
                            v1.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //2026_UTILIDAD NETA DEPARTAMENTAL AUTOS SEMINUEVOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 2026;
                            v2.Concepto = "UTILIDAD NETA DEPARTAMENTAL AUTOS SEMINUEVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 31)
                    {
                        if (version == "V1") //1027_GASTOS DEPARTAMENTALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1027;
                            v1.Concepto = "GASTOS DEPARTAMENTALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1021_INGRESOS POR VENTA DE SERVICIO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1021;
                            v2.Concepto = "INGRESOS POR VENTA DE SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 32)
                    {
                        if (version == "V1") //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1028;
                            v1.Concepto = "UTILIDAD NETA HOJALATERIA Y PINTURA";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1022_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1022;
                            v2.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 33)
                    {
                        if (version == "V1") //1029_UTILIDAD NETA SERVICIOS ADICIONALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1029;
                            v1.Concepto = "UTILIDAD NETA SERVICIOS ADICIONALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1023_GASTOS DEPARTAMENTALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1023;
                            v2.Concepto = "GASTOS DEPARTAMENTALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 34)
                    {
                        if (version == "V1") //1030_UTILIDAD NETA DEPARTAMENTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1030;
                            v1.Concepto = "UTILIDAD NETA DEPARTAMENTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1024_UTILIDAD NETA SERVICIO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1024;
                            v2.Concepto = "UTILIDAD NETA SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 35)
                    {
                        if (version == "V1") //1031_No. FACTURAS EMITIDAS SERVICIO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1031;
                            v1.Concepto = "No. FACTURAS EMITIDAS SERVICIO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1091_GASTOS DE ADMINISTRACION SERVICIO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1091;
                            v2.Concepto = "GASTOS DE ADMINISTRACION SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 36)
                    {
                        if (version == "V1") //1032_No. FACTURAS EMITIDAS HYP
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1032;
                            v1.Concepto = "No. FACTURAS EMITIDAS HYP";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //2027_UTILIDAD NETA DEPARTAMENTAL SERVICIO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 2027;
                            v2.Concepto = "UTILIDAD NETA DEPARTAMENTAL SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 37)
                    {
                        if (version == "V1") //1033_No. HORAS FACTURADAS SERVICIO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1033;
                            v1.Concepto = "No. HORAS FACTURADAS SERVICIO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1025_INGRESOS POR HOJALATERIA Y PINTURA
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1025;
                            v2.Concepto = "INGRESOS POR HOJALATERIA Y PINTURA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 38)
                    {
                        if (version == "V1") //1034_INGRESOS POR VENTA DE REFACCIONES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1034;
                            v1.Concepto = "INGRESOS POR VENTA DE REFACCIONES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1026_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1026;
                            v2.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 39)
                    {
                        if (version == "V1") //1035_VENTA DE REFACCIONES SERVICIO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1035;
                            v1.Concepto = "VENTA DE REFACCIONES SERVICIO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1027_GASTOS DEPARTAMENTALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1027;
                            v2.Concepto = "GASTOS DEPARTAMENTALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 40)
                    {
                        if (version == "V1") //1036_VENTA DE REFACCIONES H&P
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1036;
                            v1.Concepto = "VENTA DE REFACCIONES H&P";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1028_UTILIDAD NETA HOJALATERIA Y PINTURA
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1028;
                            v2.Concepto = "UTILIDAD NETA HOJALATERIA Y PINTURA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 41)
                    {
                        if (version == "V1") //1037_VENTA MAYOREO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1037;
                            v1.Concepto = "VENTA MAYOREO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1092_GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1092;
                            v2.Concepto = "GASTOS DE ADMINISTRACION HOJALATERIA Y PINTURA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 42)
                    {
                        if (version == "V1") //1038_VENTA MOSTRADOR
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1038;
                            v1.Concepto = "VENTA MOSTRADOR";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //2028_UTILIDAD NETA DEPARTAMENTAL HOJALATERIA Y PINTURA
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 2028;
                            v2.Concepto = "UTILIDAD NETA DEPARTAMENTAL HOJALATERIA Y PINTURA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 43)
                    {
                        if (version == "V1") //1039_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1039;
                            v1.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1029_UTILIDAD NETA SERVICIOS ADICIONALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1029;
                            v2.Concepto = "UTILIDAD NETA SERVICIOS ADICIONALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 44)
                    {
                        if (version == "V1") //1040_GASTOS DEPARTAMENTALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1040;
                            v1.Concepto = "GASTOS DEPARTAMENTALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1030_UTILIDAD NETA DEPARTAMENTAL
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1030;
                            v2.Concepto = "UTILIDAD NETA DEPARTAMENTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 45)
                    {
                        if (version == "V1") //1041_UTILIDAD NETA REFACCIONES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1041;
                            v1.Concepto = "UTILIDAD NETA REFACCIONES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1031_No. FACTURAS EMITIDAS SERVICIO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1031;
                            v2.Concepto = "No. FACTURAS EMITIDAS SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 46)
                    {
                        if (version == "V1") //1042_UTILIDAD NETA SERVICIOS ADICIONALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1042;
                            v1.Concepto = "UTILIDAD NETA SERVICIOS ADICIONALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1032_No. FACTURAS EMITIDAS HYP
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1032;
                            v2.Concepto = "No. FACTURAS EMITIDAS HYP";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 47)
                    {
                        if (version == "V1") //1043_UTILIDAD NETA DEPARTAMENTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1043;
                            v1.Concepto = "UTILIDAD NETA DEPARTAMENTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1033_No. HORAS FACTURADAS SERVICIO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1033;
                            v2.Concepto = "No. HORAS FACTURADAS SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 48)
                    {
                        if (version == "V1") //1044_GASTOS ADMINISTRATIVOS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1044;
                            v1.Concepto = "GASTOS ADMINISTRATIVOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1034_INGRESOS POR VENTA DE REFACCIONES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1034;
                            v2.Concepto = "INGRESOS POR VENTA DE REFACCIONES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 49)
                    {
                        if (version == "V1") //1045_OTROS INGRESOS PLANTA
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1045;
                            v1.Concepto = "OTROS INGRESOS PLANTA";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1035_VENTA DE REFACCIONES SERVICIO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1035;
                            v2.Concepto = "VENTA DE REFACCIONES SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 50)
                    {
                        if (version == "V1") //1046_UTILIDAD DE OPERACION
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1046;
                            v1.Concepto = "UTILIDAD DE OPERACION";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1036_VENTA DE REFACCIONES H&P
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1036;
                            v2.Concepto = "VENTA DE REFACCIONES H&P";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 51)
                    {
                        if (version == "V1") //1047_EBITDA
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1047;
                            v1.Concepto = "EBITDA";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1037_VENTA MAYOREO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1037;
                            v2.Concepto = "VENTA MAYOREO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 52)
                    {
                        if (version == "V1") //1048_FINANCIAMIENTO NETO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1048;
                            v1.Concepto = "FINANCIAMIENTO NETO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1038_VENTA MOSTRADOR
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1038;
                            v2.Concepto = "VENTA MOSTRADOR";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 53)
                    {
                        if (version == "V1") //1049_OTROS (GASTOS) PRODUCTOS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1049;
                            v1.Concepto = "OTROS (GASTOS) PRODUCTOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1039_UTILIDAD BRUTA DEPARTAMENTAL
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1039;
                            v2.Concepto = "UTILIDAD BRUTA DEPARTAMENTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 54)
                    {
                        if (version == "V1") //1050_GASTOS CORPORATIVOS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1050;
                            v1.Concepto = "GASTOS CORPORATIVOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1040_GASTOS DEPARTAMENTALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1040;
                            v2.Concepto = "GASTOS DEPARTAMENTALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 55)
                    {
                        if (version == "V1") //1051_UTILIDAD NETA ANTES DE FIDEICOMISO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1051;
                            v1.Concepto = "UTILIDAD NETA ANTES DE FIDEICOMISO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1042_UTILIDAD NETA SERVICIOS ADICIONALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1042;
                            v2.Concepto = "UTILIDAD NETA SERVICIOS ADICIONALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 56)
                    {
                        if (version == "V1") //1052_UTILIDAD NETA DEL FIDEICOMISO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1052;
                            v1.Concepto = "UTILIDAD NETA DEL FIDEICOMISO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1041_UTILIDAD NETA REFACCIONES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1041;
                            v2.Concepto = "UTILIDAD NETA REFACCIONES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 57)
                    {
                        if (version == "V1") //1053_UTILIDAD NETA TAXIS BAM
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1053;
                            v1.Concepto = "UTILIDAD NETA TAXIS BAM";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1087_FINANCIAMIENTO NETO REFACCIONES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1087;
                            v2.Concepto = "FINANCIAMIENTO NETO REFACCIONES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 58)
                    {
                        if (version == "V1") //1054_PARTIDAS EXTRAORDINARIAS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1054;
                            v1.Concepto = "PARTIDAS EXTRAORDINARIAS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1093_GASTOS DE ADMINISTRACION REFACCIONES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1093;
                            v2.Concepto = "GASTOS DE ADMINISTRACION REFACCIONES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 59)
                    {
                        if (version == "V1") //1055_UTILIDAD FINAL ANTES DE IMPUESTOS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1055;
                            v1.Concepto = "UTILIDAD FINAL ANTES DE IMPUESTOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //2029_UTILIDAD NETA DEPARTAMENTAL REFACCIONES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 2029;
                            v2.Concepto = "UTILIDAD NETA DEPARTAMENTAL REFACCIONES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 60)
                    {
                        if (version == "V1") //1056_I.S.R. CORRIENTE
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1056;
                            v1.Concepto = "I.S.R. CORRIENTE";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1045_OTROS INGRESOS PLANTA
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1045;
                            v2.Concepto = "OTROS INGRESOS PLANTA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 61)
                    {
                        if (version == "V1") //1057_I.S.R. DIFERIDO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1057;
                            v1.Concepto = "I.S.R. DIFERIDO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1046_UTILIDAD DE OPERACION
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1046;
                            v2.Concepto = "UTILIDAD DE OPERACION";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 62)
                    {
                        if (version == "V1") //1058_P.T.U.
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1058;
                            v1.Concepto = "P.T.U.";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1047_EBITDA
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1047;
                            v2.Concepto = "EBITDA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 63)
                    {
                        if (version == "V1") //1059_UTILIDAD NETA
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1059;
                            v1.Concepto = "UTILIDAD NETA";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1048_FINANCIAMIENTO NETO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1048;
                            v2.Concepto = "FINANCIAMIENTO NETO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 64)
                    {
                        if (version == "V1") //1060_ABSORCION DE GASTOS
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1060;
                            v1.Concepto = "ABSORCION DE GASTOS";
                            v1.Valor = celda == null ? 0 : (valor.Trim() == "" ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0)));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1049_OTROS (GASTOS) PRODUCTOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1049;
                            v2.Concepto = "OTROS (GASTOS) PRODUCTOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 65)
                    {
                        if (version == "V1") //1061_GASTOS TOTALES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1061;
                            v1.Concepto = "GASTOS TOTALES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1050_GASTOS CORPORATIVOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1050;
                            v2.Concepto = "GASTOS CORPORATIVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 66)
                    {
                        if (version == "V1") //1062_UTILIDAD BRUTA DE SERVICIO
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1062;
                            v1.Concepto = "UTILIDAD BRUTA DE SERVICIO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1051_UTILIDAD NETA ANTES DE FIDEICOMISO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1051;
                            v2.Concepto = "UTILIDAD NETA ANTES DE FIDEICOMISO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 67)
                    {
                        if (version == "V1") //1063_UTILIDAD BRUTA DE REFACCIONES
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1063;
                            v1.Concepto = "UTILIDAD BRUTA DE REFACCIONES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1052_UTILIDAD NETA DEL FIDEICOMISO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1052;
                            v2.Concepto = "UTILIDAD NETA DEL FIDEICOMISO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 68)
                    {
                        if (version == "V1") //1064_% DE ABSORCIÓN
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1064;
                            v1.Concepto = "% DE ABSORCIÓN";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1053_UTILIDAD NETA TAXIS BAM
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1053;
                            v2.Concepto = "UTILIDAD NETA TAXIS BAM";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 69)
                    {
                        if (version == "V1") //1065_% DE GASTOS SOBRE UTILIDAD BRUTA TOTAL
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 5;
                            v1.IdConcepto = 1065;
                            v1.Concepto = "% DE GASTOS SOBRE UTILIDAD BRUTA TOTAL";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else //1054_PARTIDAS EXTRAORDINARIAS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1054;
                            v2.Concepto = "PARTIDAS EXTRAORDINARIAS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 70)
                    {
                        if (version == "V2") //1055_UTILIDAD FINAL ANTES DE IMPUESTOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1055;
                            v2.Concepto = "UTILIDAD FINAL ANTES DE IMPUESTOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 71)
                    {
                        if (version == "V2") //1056_I.S.R. CORRIENTE
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1056;
                            v2.Concepto = "I.S.R. CORRIENTE";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 72)
                    {
                        if (version == "V2") //1057_I.S.R. DIFERIDO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1057;
                            v2.Concepto = "I.S.R. DIFERIDO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 73)
                    {
                        if (version == "V2") //1058_P.T.U.
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1058;
                            v2.Concepto = "P.T.U.";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 74)
                    {
                        if (version == "V2") //1059_UTILIDAD NETA
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1059;
                            v2.Concepto = "UTILIDAD NETA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 75)
                    {
                        if (version == "V2") //1060_ABSORCION DE GASTOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1060;
                            v2.Concepto = "ABSORCION DE GASTOS";
                            v2.Valor = celda == null ? 0 : (valor.Trim() == "" ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0)));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 76)
                    {
                        if (version == "V2") //1061_GASTOS TOTALES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1061;
                            v2.Concepto = "GASTOS TOTALES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 77)
                    {
                        if (version == "V2") //1044_GASTOS ADMINISTRATIVOS
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1044;
                            v2.Concepto = "GASTOS ADMINISTRATIVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 78)
                    {
                        if (version == "V2") //1062_UTILIDAD BRUTA DE SERVICIO
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1062;
                            v2.Concepto = "UTILIDAD BRUTA DE SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 79)
                    {
                        if (version == "V2") //1063_UTILIDAD BRUTA DE REFACCIONES
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1063;
                            v2.Concepto = "UTILIDAD BRUTA DE REFACCIONES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 80)
                    {
                        if (version == "V2") //1064_% DE ABSORCIÓN
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1064;
                            v2.Concepto = "% DE ABSORCIÓN";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 81)
                    {
                        if (version == "V2") //1065_% DE GASTOS SOBRE UTILIDAD BRUTA TOTAL
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 5;
                            v2.IdConcepto = 1065;
                            v2.Concepto = "% DE GASTOS SOBRE UTILIDAD BRUTA TOTAL";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                }

                c++;

                if (r == hoja.LastRow)
                    r  = 1;

                if (c == hoja.LastColumn + 2)
                {
                    c = 1;
                    r++;
                }
            }

            //for (int i = 1; i < hoja.LastRow; i++)
            //{
            //    for (int j = 1; j < hoja.LastColumn; j++)
            //    {
            //        if (i == 1)
            //        {
                        
            //        }
            //    }
            //}
        }

        public void LeePestanaBG(DVAExcel.HojaDeTrabajo hoja)
        {
            string valor = "";
            int r = 1;
            int c = 1;

            ProcesoResultadoMensualExcel v1;
            ProcesoResultadoMensualExtralibrosExcel v2;

            foreach (object celda in hoja.Celdas)
            {
                if (celda != null)
                    valor = celda.ToString();

                if ((c == (mes + 1)) && r >= 4)
                {
                    if (r == 8) //41_Efectivo e Inversiones
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 41;
                            v1.Concepto = "Efectivo e Inversiones";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 41;
                            v2.Concepto = "Efectivo e Inversiones";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 10) //42_Documentos por Cobrar Unidades
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 42;
                            v1.Concepto = "Documentos por Cobrar Unidades";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 42;
                            v2.Concepto = "Documentos por Cobrar Unidades";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 11) //43_Cuentas por Cobrar Unidades
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 43;
                            v1.Concepto = "Cuentas por Cobrar Unidades";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 43;
                            v2.Concepto = "Cuentas por Cobrar Unidades";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 12) //44_Cuentas por Cobrar Refacciones
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 44;
                            v1.Concepto = "Cuentas por Cobrar Refacciones";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 44;
                            v2.Concepto = "Cuentas por Cobrar Refacciones";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 13) //45_Cuentas por Cobrar Servicio
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 45;
                            v1.Concepto = "Cuentas por Cobrar Servicio";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 45;
                            v2.Concepto = "Cuentas por Cobrar Servicio";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 14) //46_Reserva Incobrables
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 46;
                            v1.Concepto = "Reserva Incobrables";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 46;
                            v2.Concepto = "Reserva Incobrables";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 17) //47_Planta Activo
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 47;
                            v1.Concepto = "Planta Activo";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 47;
                            v2.Concepto = "Planta Activo";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 18) //48_Deudores Diversos y Funcionarios y Empleados
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 48;
                            v1.Concepto = "Deudores Diversos y Funcionarios y Empleados";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 48;
                            v2.Concepto = "Deudores Diversos y Funcionarios y Empleados";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }                    
                    else if (r == 19) //49_Partes Relacionadas Activo
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 49;
                            v1.Concepto = "Partes Relacionadas Activo";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 49;
                            v2.Concepto = "Partes Relacionadas Activo";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 22) //50_Impuestos por Recuperar
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 50;
                            v1.Concepto = "Impuestos por Recuperar";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 50;
                            v2.Concepto = "Impuestos por Recuperar";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 23) //51_Pagos Anticipados
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 51;
                            v1.Concepto = "Pagos Anticipados";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 51;
                            v2.Concepto = "Pagos Anticipados";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 25) //52_Inventario Autos Nuevos
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 52;
                            v1.Concepto = "Inventario Autos Nuevos";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 52;
                            v2.Concepto = "Inventario Autos Nuevos";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 26) //53_Inventario Autos Seminuevos
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 53;
                            v1.Concepto = "Inventario Autos Seminuevos";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 53;
                            v2.Concepto = "Inventario Autos Seminuevos";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 27) //54_Inventario Refacciones y Accesorios
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 54;
                            v1.Concepto = "Inventario Refacciones y Accesorios";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 54;
                            v2.Concepto = "Inventario Refacciones y Accesorios";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 28) //55_Inventarios Otros y Proceso
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 55;
                            v1.Concepto = "Inventarios Otros y Proceso";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 55;
                            v2.Concepto = "Inventarios Otros y Proceso";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 33) //56_Inversiones Permanentes en Acciones
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 56;
                            v1.Concepto = "Inversiones Permanentes en Acciones";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 56;
                            v2.Concepto = "Inversiones Permanentes en Acciones";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 36) //58_Terreno y Edificio
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 58;
                            v1.Concepto = "Terreno y Edificio";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 58;
                            v2.Concepto = "Terreno y Edificio";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 37) //59_Mobiliario y Equipo
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 59;
                            v1.Concepto = "Mobiliario y Equipo";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 59;
                            v2.Concepto = "Mobiliario y Equipo";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 38) //60_Equipo de Transporte
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 60;
                            v1.Concepto = "Equipo de Transporte";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 60;
                            v2.Concepto = "Equipo de Transporte";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 39) //61_Depreciación Acumulada
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 61;
                            v1.Concepto = "Depreciación Acumulada";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 61;
                            v2.Concepto = "Depreciación Acumulada";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 42) //62_Otros Activos
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 62;
                            v1.Concepto = "Otros Activos";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 62;
                            v2.Concepto = "Otros Activos";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 43) //63_Mejoras Inmuebles Arrendados
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 63;
                            v1.Concepto = "Mejoras Inmuebles Arrendados";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 63;
                            v2.Concepto = "Mejoras Inmuebles Arrendados";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 44) //64_Neto de Actualizaciones
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 64;
                            v1.Concepto = "Neto de Actualizaciones";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 64;
                            v2.Concepto = "Neto de Actualizaciones";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 52) //66_Planta Pasivo
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 66;
                            v1.Concepto = "Planta Pasivo";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 66;
                            v2.Concepto = "Planta Pasivo";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 53) //67_Impuestos por Pagar
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 67;
                            v1.Concepto = "Impuestos por Pagar";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 67;
                            v2.Concepto = "Impuestos por Pagar";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 54) //68_Anticipos de Clientes
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 68;
                            v1.Concepto = "Anticipos de Clientes";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 68;
                            v2.Concepto = "Anticipos de Clientes";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 56) //70_Proveedores
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 70;
                            v1.Concepto = "Proveedores";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 70;
                            v2.Concepto = "Proveedores";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 57) //71_Acreedores Diversos
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 71;
                            v1.Concepto = "Acreedores Diversos";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 71;
                            v2.Concepto = "Acreedores Diversos";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 58) //72_P.T.U.
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 72;
                            v1.Concepto = "P.T.U.";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 72;
                            v2.Concepto = "P.T.U.";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 59) //73_Partes Relacionadas Pasivo
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 73;
                            v1.Concepto = "Partes Relacionadas Pasivo";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 73;
                            v2.Concepto = "Partes Relacionadas Pasivo";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 60) //74_Documentos por Pagar
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 74;
                            v1.Concepto = "Documentos por Pagar";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 74;
                            v2.Concepto = "Documentos por Pagar";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 61) //75_Otras Provisiones y Cuentas por Pagar
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 75;
                            v1.Concepto = "Otras Provisiones y Cuentas por Pagar";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 75;
                            v2.Concepto = "Otras Provisiones y Cuentas por Pagar";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 62) //76_I.S.R. Diferido
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 76;
                            v1.Concepto = "I.S.R. Diferido";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 76;
                            v2.Concepto = "I.S.R. Diferido";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 67) //79_Intereses
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 79;
                            v1.Concepto = "Intereses";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 79;
                            v2.Concepto = "Intereses";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 68) //80_Prima de Antigüedad
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 80;
                            v1.Concepto = "Prima de Antigüedad";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 80;
                            v2.Concepto = "Prima de Antigüedad";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 76) //81_Capital Social
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 81;
                            v1.Concepto = "Capital Social";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 81;
                            v2.Concepto = "Capital Social";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 77) //82_Reserva Legal
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 82;
                            v1.Concepto = "Reserva Legal";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 82;
                            v2.Concepto = "Reserva Legal";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 78) //83_Resultado de Ejercicios Anteriores
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 83;
                            v1.Concepto = "Resultado de Ejercicios Anteriores";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 83;
                            v2.Concepto = "Resultado de Ejercicios Anteriores";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 79) //84_Resultado del Ejercicio
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 84;
                            v1.Concepto = "Resultado del Ejercicio";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 84;
                            v2.Concepto = "Resultado del Ejercicio";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 80) //85_Actualización del Capital Contable
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdReporte = 4;
                            v1.IdConcepto = 85;
                            v1.Concepto = "Actualización del Capital Contable";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v1);

                            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdReporte = 4;
                            v2.IdConcepto = 85;
                            v2.Concepto = "Actualización del Capital Contable";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                            _db.Insert(3665, 0, v2);

                            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                }

                c++;

                if (r == hoja.LastRow)
                    r = 1;

                if (c == hoja.LastColumn + 2)
                {
                    c = 1;
                    r++;
                }
            }
        }

        public void LeerPestanaVP(DVAExcel.HojaDeTrabajo hoja) 
        {
            string valor = "";
            int r = 1;
            int c = 1;

            ProcesoResultadoMensualExcel v1;
            ProcesoResultadoMensualExtralibrosExcel v2;

            foreach (object celda in hoja.Celdas)
            {

                if (celda != null)
                    valor = celda.ToString();

                if ((c == (mes + 1)) && r >= 4)
                {
                    if (r == 4)  //TOTAL DE UNIDADES NUEVAS VENDIDAS MES
                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();
                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 86;
                            v1.Concepto = "TOTAL DE UNIDADES NUEVAS VENDIDAS MES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);

                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 86;
                            v2.Concepto = "TOTAL DE UNIDADES NUEVAS VENDIDAS MES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                        }
                    }
                    else if (r == 5) //TOTAL DE UNIDADES NUEVAS VENDIDAS ACUMULADO

                    {
                        if (version == "V1")
                        {
                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 87;
                            v1.Concepto = "TOTAL DE UNIDADES NUEVAS VENDIDAS ACUMULADO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 87;
                            v2.Concepto = "TOTAL DE UNIDADES NUEVAS VENDIDAS ACUMULADO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 7)//UNIDADES VENDIDAS AGENCIA MENUDEO MES

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 88;
                            v1.Concepto = "UNIDADES VENDIDAS AGENCIA MENUDEO MES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 88;
                            v2.Concepto = "UNIDADES VENDIDAS AGENCIA MENUDEO MES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 8)//UTILIDAD BRUTA MENUDEO

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 89;
                            v1.Concepto = "UTILIDAD BRUTA MENUDEO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 89;
                            v2.Concepto = "UTILIDAD BRUTA MENUDEO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 9)//UTILIDAD BRUTA PROMEDIO UNIDAD NUEVA VENDIDA

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 90;
                            v1.Concepto = "UTILIDAD BRUTA PROMEDIO UNIDAD NUEVA VENDIDA";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 90;
                            v2.Concepto = "UTILIDAD BRUTA PROMEDIO UNIDAD NUEVA VENDIDA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 11)//UNIDADES VENDIDAS AGENCIA FLOTILLA MES

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 91;
                            v1.Concepto = "UNIDADES VENDIDAS AGENCIA FLOTILLA MES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 91;
                            v2.Concepto = "UNIDADES VENDIDAS AGENCIA FLOTILLA MES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 12)//UTILIDAD BRUTA FLOTILLAS

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 92;
                            v1.Concepto = "UTILIDAD BRUTA FLOTILLAS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 92;
                            v2.Concepto = "UTILIDAD BRUTA FLOTILLAS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 13)//U. B. PROMEDIO UNIDAD NUEVA VENDIDA FLOTILLA

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 93;
                            v1.Concepto = "U. B. PROMEDIO UNIDAD NUEVA VENDIDA FLOTILLA";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 93;
                            v2.Concepto = "U. B. PROMEDIO UNIDAD NUEVA VENDIDA FLOTILLA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 15)//UNIDADES VENDIDAS POR INTERNET

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 94;
                            v1.Concepto = "UNIDADES VENDIDAS POR INTERNET";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 94;
                            v2.Concepto = "UNIDADES VENDIDAS POR INTERNET";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 16)//UTILIDAD BRUTA FLOTILLAS (UTILIDAD BRUTA VENDIDAS POR INTERNET - en tabla)

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 95;
                            v1.Concepto = "UTILIDAD BRUTA VENDIDAS POR INTERNET";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 95;
                            v2.Concepto = "UTILIDAD BRUTA VENDIDAS POR INTERNET";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 17)//U. B. PROMEDIO UNIDAD NUEVA VENDIDA FLOTILLA (U. B. PROMEDIO UNIDAD NUEVA VENDIDA POR INTERNET- en tabla)

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 96;
                            v1.Concepto = "U. B. PROMEDIO UNIDAD NUEVA VENDIDA POR INTERNET";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 96;
                            v2.Concepto = "U. B. PROMEDIO UNIDAD NUEVA VENDIDA POR INTERNET";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 19)//UNIDADES VENDIDAS A AUTOFIN MES

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 97;
                            v1.Concepto = "UNIDADES VENDIDAS A AUTOFIN MES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 97;
                            v2.Concepto = "UNIDADES VENDIDAS A AUTOFIN MES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 20)//UTILIDAD BRUTA AUTOFIN

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 98;
                            v1.Concepto = "UTILIDAD BRUTA AUTOFIN";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 98;
                            v2.Concepto = "UTILIDAD BRUTA AUTOFIN";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 21)//U. B. PROMEDIO UNIDAD NUEVA VENDIDA AUTOFIN

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 99;
                            v1.Concepto = "U. B. PROMEDIO UNIDAD NUEVA VENDIDA AUTOFIN";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 99;
                            v2.Concepto = "U. B. PROMEDIO UNIDAD NUEVA VENDIDA AUTOFIN";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 23)//UNIDADES VENDIDAS TAXIS BAM MES

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 100;
                            v1.Concepto = "UNIDADES VENDIDAS TAXIS BAM MES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 100;
                            v2.Concepto = "UNIDADES VENDIDAS TAXIS BAM MES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }


                    else if (r == 24)//UTILIDAD BRUTA TAXIS BAM

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 101;
                            v1.Concepto = "UTILIDAD BRUTA TAXIS BAM";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 101;
                            v2.Concepto = "UTILIDAD BRUTA TAXIS BAM";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 25)//U. B. PROMEDIO UNIDAD NUEVA VENDIDA TAXIS BAM

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 102;
                            v1.Concepto = "U. B. PROMEDIO UNIDAD NUEVA VENDIDA TAXIS BAM";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 102;
                            v2.Concepto = "U. B. PROMEDIO UNIDAD NUEVA VENDIDA TAXIS BAM";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 27)//PARTICIPACION AUTOFIN MES

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 103;
                            v1.Concepto = "PARTICIPACION AUTOFIN MES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 103;
                            v2.Concepto = "PARTICIPACION AUTOFIN MES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 28)//PARTICIPACION AUTOFIN ACUMULADO

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 104;
                            v1.Concepto = "PARTICIPACION AUTOFIN ACUMULADO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 104;
                            v2.Concepto = "PARTICIPACION AUTOFIN ACUMULADO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 30)//SEMINUEVOS VENDIDOS MES

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 105;
                            v1.Concepto = "SEMINUEVOS VENDIDOS MES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 105;
                            v2.Concepto = "SEMINUEVOS VENDIDOS MES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 31)//SEMINUEVOS VENDIDOS ACUMULADO

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 106;
                            v1.Concepto = "SEMINUEVOS VENDIDOS ACUMULADO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 106;
                            v2.Concepto = "SEMINUEVOS VENDIDOS ACUMULADO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 33)//SEMINUEVOS VENDIDOS ACUMULADO

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 107;
                            v1.Concepto = "UTILIDAD BRUTA UNIDADES SEMINUEVAS MES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 107;
                            v2.Concepto = "UTILIDAD BRUTA UNIDADES SEMINUEVAS MES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 34)//U.B. PROM. UNIDADES SEMINUEVAS VENDIDAS MES

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 108;
                            v1.Concepto = "U.B. PROM. UNIDADES SEMINUEVAS VENDIDAS MES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 108;
                            v2.Concepto = "U.B. PROM. UNIDADES SEMINUEVAS VENDIDAS MES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 36)//UTILIDAD BRUTA SERVICIO

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 109;
                            v1.Concepto = "UTILIDAD BRUTA SERVICIO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 109;
                            v2.Concepto = "UTILIDAD BRUTA SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 38)//UTILIDAD BRUTA HOJALATERIA Y PINTURA

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 110;
                            v1.Concepto = "UTILIDAD BRUTA HOJALATERIA Y PINTURA";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 110;
                            v2.Concepto = "UTILIDAD BRUTA HOJALATERIA Y PINTURA";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }

                    else if (r == 40)//UTILIDAD BRUTA REFACCIONES

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 111;
                            v1.Concepto = "UTILIDAD BRUTA REFACCIONES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 111;
                            v2.Concepto = "UTILIDAD BRUTA REFACCIONES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 42)//UTILIDAD NETA NUEVOS

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 112;
                            v1.Concepto = "UTILIDAD NETA NUEVOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 112;
                            v2.Concepto = "UTILIDAD NETA NUEVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 44)//UTILIDAD NETA SEMINUEVOS

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 113;
                            v1.Concepto = "UTILIDAD NETA SEMINUEVOS";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 113;
                            v2.Concepto = "UTILIDAD NETA SEMINUEVOS";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 46)//UTILIDAD NETA SERVICIO

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 114;
                            v1.Concepto = "UTILIDAD NETA SERVICIO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 114;
                            v2.Concepto = "UTILIDAD NETA SERVICIO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 48)//UTILIDAD NETA REFACCIONES

                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 115;
                            v1.Concepto = "UTILIDAD NETA REFACCIONES";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 115;
                            v2.Concepto = "UTILIDAD NETA REFACCIONES";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                    else if (r == 50)//UTILIDAD ANTES DE FIDEICOMISO


                    {
                        if (version == "V1")
                        {

                            v1 = new ProcesoResultadoMensualExcel();

                            v1.IdAgencia = idAgencia;
                            v1.Anio = anio;
                            v1.IdMes = mes;
                            v1.IdConcepto = 116;
                            v1.Concepto = "UTILIDAD ANTES DE FIDEICOMISO";
                            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v1.IdReporte = 9;

                            _db.Insert(32621, 0, v1);

                            Console.Write("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                        }
                        else
                        {
                            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                            v2.IdAgencia = idAgencia;
                            v2.Anio = anio;
                            v2.IdMes = mes;
                            v2.IdConcepto = 116;
                            v2.Concepto = "UTILIDAD ANTES DE FIDEICOMISO";
                            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor) * 100, 0));
                            v2.IdReporte = 9;

                            _db.Insert(32621, 0, v2);

                            Console.Write("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);

                        }

                    }
                }

                c++;

                //if (r == hoja.LastRow)
                //  r = 1;

                if (c == hoja.LastColumn + 2)
                {
                    c = 1;
                    r++;
                }
            }
        }

        public void LeePestanaCuentas(DVAExcel.HojaDeTrabajo hoja)
        {
            string valor = "";
            int r = 3;
            int c = 1;    

            ProcesoResultadoMensualExcel v1;
            ProcesoResultadoMensualExtralibrosExcel v2;
            var celdaCoordenadas = hoja.Celdas[3,3];
            var getValuesRow = hoja.GetStrRow(3);
            var testt = hoja.GetRowValues(3);
            var lastROw = hoja.LastRow;
            var lastColumn_ = hoja.LastColumn;

            var columnaMes = mes + 2;

            var concepto1 = hoja.Celdas[3 ,columnaMes]; //Cuentas por cobrar unidades Nuevas (bancos, cartera)
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 709;
                v1.Concepto = "Cuentas por Cobrar Unidades Nuevas (Bancos, cartera 1130)";
                if (concepto1.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto1), 0));
                }
               

                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 709;
                v2.Concepto = "Cuentas por Cobrar Unidades Nuevas (Bancos, cartera 1130)";
              

                if (concepto1.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto1), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            var concepto2 = hoja.Celdas[4, columnaMes]; //Cuentas por Cobrar Financieras de Marca
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 710;
                v1.Concepto = "Cuentas por Cobrar Financieras de Marca";
                if (concepto2.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto2), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 710;
                v2.Concepto = "Cuentas por Cobrar Financieras de Marca";
                if (concepto2.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto2), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            var concepto3 = hoja.Celdas[5, columnaMes]; //Cuentas por Cobrar Financieras de Marca
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Contratos en Transito";
                if (concepto3.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto3), 0));
                }
                //_db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Contratos en Transito";
                if (concepto3.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto3), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            var concepto4 = hoja.Celdas[7, columnaMes]; //Cuentas por Cobrar Financieras de Marca
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 711;
                v1.Concepto = "Cuentas por Cobrar Unidades Nuevas";
                if (concepto4.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto4), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 711;
                v2.Concepto = "Cuentas por Cobrar Unidades Nuevas";
                if (concepto4.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto4), 0));
                }

                 _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            var concepto5 = hoja.Celdas[9, columnaMes]; //Cuentas por Cobrar Financieras de Marca
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 712;
                v1.Concepto = "Cuentas por Cobrar Garantias";
                if (concepto5.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto5), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 712;
                v2.Concepto = "Cuentas por Cobrar Garantias";
                if (concepto5.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto5), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto6 = hoja.Celdas[12, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 713;
                v1.Concepto = "Venta Mano de Obra Taller";
                if (concepto6.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto6), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 713;
                v2.Concepto = "Venta Mano de Obra Taller";
                if (concepto6.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto6), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto7 = hoja.Celdas[13, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 714;
                v1.Concepto = "Devolución Venta Mano de Obra Taller";
                if (concepto7.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto7), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 714;
                v2.Concepto = "Devolución Venta Mano de Obra Taller";
                if (concepto7.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto7), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto8 = hoja.Celdas[14, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 715;
                v1.Concepto = "Costo de Venta Mano de Obra Taller";
                if (concepto8.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto8), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 715;
                v2.Concepto = "Costo de Venta Mano de Obra Taller";
                if (concepto8.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto8), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto9 = hoja.Celdas[15, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 716;
                v1.Concepto = "Recuperacion de Mano de Obra";
                if (concepto9.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto9), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 716;
                v2.Concepto = "Recuperacion de Mano de Obra";
                if (concepto9.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto9), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto10 = hoja.Celdas[16, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 717;
                v1.Concepto = "Costo M.O. Mecanica Proceso";
                if (concepto10.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto10), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 717;
                v2.Concepto = "Costo M.O. Mecanica Proceso";
                if (concepto10.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto10), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto11 = hoja.Celdas[17, columnaMes]; //Utilidad bruta de mano de obra 
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Mano de Obra Taller";
                if (concepto11.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto11), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Mano de Obra Taller";
                if (concepto11.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto11), 0));
                }

                //_db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto12 = hoja.Celdas[19, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 718;
                v1.Concepto = "Venta Mano de Obra Garantias";
                if (concepto12.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto12), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 718;
                v2.Concepto = "Venta Mano de Obra Garantias";
                if (concepto12.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto12), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto13 = hoja.Celdas[20, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 719;
                v1.Concepto = "Devolución Venta Mano de Obra Garantias";
                if (concepto13.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto13), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 719;
                v2.Concepto = "Devolución Venta Mano de Obra Garantias";
                if (concepto13.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto13), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto14 = hoja.Celdas[21, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 720;
                v1.Concepto = "Costo de Venta Mano de Obra Garantias";
                if (concepto14.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto14), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 720;
                v2.Concepto = "Costo de Venta Mano de Obra Garantias";
                if (concepto14.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto14), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto15 = hoja.Celdas[22, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Mano de Obra Garantias";
                if (concepto15.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto15), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Mano de Obra Garantias";
                if (concepto15.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto15), 0));
                }

                //_db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto16 = hoja.Celdas[24, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 721;
                v1.Concepto = "Ventas T.O.T.";
                if (concepto16.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto16), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 721;
                v2.Concepto = "Ventas T.O.T.";
                if (concepto16.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto16), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto17 = hoja.Celdas[25, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 722;
                v1.Concepto = "Devolución T.O.T.";
                if (concepto17.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto17), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 722;
                v2.Concepto = "Devolución T.O.T.";
                if (concepto17.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto17), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto18 = hoja.Celdas[26, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 723;
                v1.Concepto = "Costo de Vetas T.O.T.";
                if (concepto18.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto18), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 723;
                v2.Concepto = "Costo de Vetas T.O.T.";
                if (concepto18.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto18), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto19 = hoja.Celdas[26, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta T.O.T.";
                if (concepto19.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto19), 0));
                }
                //_db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta T.O.T.";
                if (concepto19.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto19), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto20 = hoja.Celdas[29, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 724;
                v1.Concepto = "Venta Materiales Diversos.";
                if (concepto20.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto20), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 724;
                v2.Concepto = "Venta Materiales Diversos.";
                if (concepto20.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto20), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto21 = hoja.Celdas[30, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 725;
                v1.Concepto = "Devolución Venta Materiales Diversos.";
                if (concepto21.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto21), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 725;
                v2.Concepto = "Devolución Venta Materiales Diversos.";
                if (concepto21.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto21), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto22 = hoja.Celdas[31, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 726;
                v1.Concepto = "Costo de Venta Materiales Diversos.";
                if (concepto22.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto22), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 726;
                v2.Concepto = "Costo de Venta Materiales Diversos.";
                if (concepto22.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto22), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto23 = hoja.Celdas[32, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Materiales Diversos.";
                if (concepto23.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto23), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Materiales Diversos.";
                if (concepto23.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto23), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto24 = hoja.Celdas[35, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 727;
                v1.Concepto = "Venta Refacciones Mayoreo.";
                if (concepto24.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto24), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 727;
                v2.Concepto = "Venta Refacciones Mayoreo.";
                if (concepto24.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto24), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto25 = hoja.Celdas[36, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 728;
                v1.Concepto = "Devolución Venta Refacciones Mayoreo.";
                if (concepto25.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto25), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 728;
                v2.Concepto = "Devolución Venta Refacciones Mayoreo.";
                if (concepto25.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto25), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }


            //--------------------------------------------------------------------------------------------

            var concepto26 = hoja.Celdas[37, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 729;
                v1.Concepto = "Costo de Venta Refacciones Mayoreo.";
                if (concepto26.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto26), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 729;
                v2.Concepto = "Costo de Venta Refacciones Mayoreo.";
                if (concepto26.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto26), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }


            //--------------------------------------------------------------------------------------------

            var concepto27 = hoja.Celdas[38, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 730;
                v1.Concepto = "Ingresos por Bonificaciones Refacciones";
                if (concepto27.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto27), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 730;
                v2.Concepto = "Ingresos por Bonificaciones Refacciones";
                if (concepto27.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto27), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto28 = hoja.Celdas[39, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Refacciones Mayoreo.";
                if (concepto28.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto28), 0));
                }
                //_db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 730;
                v2.Concepto = "Utilidad Bruta Refacciones Mayoreo.";
                if (concepto28.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto28), 0));
                }

                //_db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto29 = hoja.Celdas[41, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 731;
                v1.Concepto = "Venta Refacciones Mostrador.";
                if (concepto29.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto29), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 731;
                v2.Concepto = "Venta Refacciones Mostrador.";
                if (concepto29.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto29), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto30 = hoja.Celdas[42, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 732;
                v1.Concepto = "Devolución Venta Refacciones Mostrador.";
                if (concepto30.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto30), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 732;
                v2.Concepto = "Devolución Venta Refacciones Mostrador.";
                if (concepto30.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto30), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto31 = hoja.Celdas[43, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 733;
                v1.Concepto = "Costo de Venta Refacciones Mostrador.";
                if (concepto31.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto31), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 733;
                v2.Concepto = "Costo de Venta Refacciones Mostrador.";
                if (concepto31.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto31), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto32 = hoja.Celdas[44, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Refacciones Mostrador.";
                if (concepto32.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto32), 0));
                }
                //_db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Refacciones Mostrador.";
                if (concepto32.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto32), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto33 = hoja.Celdas[46, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 734;
                v1.Concepto = "Venta Refacciones Taller.";
                if (concepto33.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto33), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 734;
                v2.Concepto = "Venta Refacciones Taller.";
                if (concepto33.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto33), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto34 = hoja.Celdas[47, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 735;
                v1.Concepto = "Devolución Venta Refacciones Taller.";
                if (concepto34.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto34), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 735;
                v2.Concepto = "Devolución Venta Refacciones Taller.";
                if (concepto34.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto34), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto35 = hoja.Celdas[48, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 736;
                v1.Concepto = "Costo de Venta Refacciones Taller.";
                if (concepto35.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto35), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 736;
                v2.Concepto = "Costo de Venta Refacciones Taller.";
                if (concepto35.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto35), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto36 = hoja.Celdas[49, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Refacciones Taller.";
                if (concepto36.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto36), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Refacciones Taller.";
                if (concepto36.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto36), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto37 = hoja.Celdas[51, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 737;
                v1.Concepto = "Venta Refacciones Cía. de Seguro.";
                if (concepto37.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto37), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 737;
                v2.Concepto = "Venta Refacciones Cía. de Seguro.";
                if (concepto37.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto37), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto38 = hoja.Celdas[52, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 738;
                v1.Concepto = "Devolución Venta Refacciones Cía. de Seguro.";
                if (concepto38.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto38), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 738;
                v2.Concepto = "Devolución Venta Refacciones Cía. de Seguro.";
                if (concepto38.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto38), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto39 = hoja.Celdas[53, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 739;
                v1.Concepto = "Costo de Venta Refacciones Cía. de Seguro.";
                if (concepto39.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto39), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 739;
                v2.Concepto = "Costo de Venta Refacciones Cía. de Seguro.";
                if (concepto39.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto39), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto40 = hoja.Celdas[54, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Refacciones Cía. de Seguro.";
                if (concepto40.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto40), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 7390;
                v2.Concepto = "Utilidad Bruta Refacciones Cía. de Seguro.";
                if (concepto40.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto40), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto41 = hoja.Celdas[56, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 740;
                v1.Concepto = "Venta Refacciones Garantía.";
                if (concepto41.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto41), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 740;
                v2.Concepto = "Venta Refacciones Garantía.";
                if (concepto41.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto41), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto42 = hoja.Celdas[57, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 741;
                v1.Concepto = "Devolución Venta Refacciones Garantía.";
                if (concepto42.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto42), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 741;
                v2.Concepto = "Devolución Venta Refacciones Garantía.";
                if (concepto42.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto42), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto43 = hoja.Celdas[58, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 742;
                v1.Concepto = "Costo de Venta Refacciones Garantía.";
                if (concepto43.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto43), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 742;
                v2.Concepto = "Costo de Venta Refacciones Garantía.";
                if (concepto43.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto43), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto44 = hoja.Celdas[59, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Refacciones Garantía.";
                if (concepto44.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto44), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Refacciones Garantía.";
                if (concepto44.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto44), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto45 = hoja.Celdas[61, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 743;
                v1.Concepto = "Venta Refacciones Otras Mercancias.";
                if (concepto45.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto45), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 743;
                v2.Concepto = "Venta Refacciones Otras Mercancias.";
                if (concepto45.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto45), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto46 = hoja.Celdas[62, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 744;
                v1.Concepto = "Devolución Venta Refacciones Otras Mercancias.";
                if (concepto46.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto46), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 744;
                v2.Concepto = "Devolución Venta Refacciones Otras Mercancias.";
                if (concepto46.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto46), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto47 = hoja.Celdas[63, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 744;
                v1.Concepto = "Costo de Venta Refacciones Otras Mercancias.";
                if (concepto47.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto47), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 744;
                v2.Concepto = "Costo de Venta Refacciones Otras Mercancias.";
                if (concepto47.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto47), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto48 = hoja.Celdas[64, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Refacciones Otras Mercancias.";
                if (concepto48.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto48), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Refacciones Otras Mercancias.";
                if (concepto48.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto48), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto49 = hoja.Celdas[66, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 746;
                v1.Concepto = "Venta Refacciones Accesorios.";
                if (concepto49.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto49), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 746;
                v2.Concepto = "Venta Refacciones Accesorios.";
                if (concepto49.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto49), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto50 = hoja.Celdas[67, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 747;
                v1.Concepto = "Devolución Venta Refacciones Accesorios.";
                if (concepto50.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto50), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 747;
                v2.Concepto = "Devolución Venta Refacciones Accesorios.";
                if (concepto50.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto50), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto51 = hoja.Celdas[68, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 747;
                v1.Concepto = "Costo de Venta Refacciones Accesorios.";
                if (concepto51.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto51), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 747;
                v2.Concepto = "Costo de Venta Refacciones Accesorios.";
                if (concepto51.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto51), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto52 = hoja.Celdas[69, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Costo de Venta Refacciones Accesorios.";
                if (concepto52.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto52), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Costo de Venta Refacciones Accesorios.";
                if (concepto52.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto52), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto53 = hoja.Celdas[72, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 749;
                v1.Concepto = "Venta Hojalateria y Pintura";
                if (concepto53.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto53), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 749;
                v2.Concepto = "Venta Hojalateria y Pintura";
                if (concepto53.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto53), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto54 = hoja.Celdas[73, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 750;
                v1.Concepto = "Devolución Venta Hojalateria y Pintura";
                if (concepto54.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto54), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 750;
                v2.Concepto = "Devolución Venta Hojalateria y Pintura";
                if (concepto54.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto54), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto55 = hoja.Celdas[74, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 751;
                v1.Concepto = "Costo de Venta Hojalateria y Pintura";
                if (concepto55.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto55), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 751;
                v2.Concepto = "Costo de Venta Hojalateria y Pintura";
                if (concepto55.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto55), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto56 = hoja.Celdas[75, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 752;
                v1.Concepto = "Costo de Venta Mano de Obra Hojalateria y Pintura Proceso";
                if (concepto56.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto56), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 752;
                v2.Concepto = "Costo de Venta Mano de Obra Hojalateria y Pintura Proceso";
                if (concepto56.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto56), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto57 = hoja.Celdas[76, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 753;
                v1.Concepto = "Recuperación Mano de Obra Hojalateria y Pintura";
                if (concepto57.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto57), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 753;
                v2.Concepto = "Recuperación Mano de Obra Hojalateria y Pintura";
                if (concepto57.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto57), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto58 = hoja.Celdas[77, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Hojalateria y Pintura.";
                if (concepto57.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto57), 0));
                }
                //_db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Hojalateria y Pintura.";
                if (concepto57.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto57), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }


            //--------------------------------------------------------------------------------------------

            var concepto59 = hoja.Celdas[79, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 754;
                v1.Concepto = "Venta Materiales Diversos Hojalateria y Pintura";
                if (concepto59.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto59), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 754;
                v2.Concepto = "Venta Materiales Diversos Hojalateria y Pintura";
                if (concepto59.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto59), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto60 = hoja.Celdas[80, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 755;
                v1.Concepto = "Devolución Venta Materiales Diversos Hojalateria y Pintura";
                if (concepto60.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto60), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 755;
                v2.Concepto = "Devolución Venta Materiales Diversos Hojalateria y Pintura";
                if (concepto60.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto60), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto61 = hoja.Celdas[81, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 755;
                v1.Concepto = "Recuperación Mano de Obra Hojalateria y Pintura";
                if (concepto61.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto61), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 755;
                v2.Concepto = "Recuperación Mano de Obra Hojalateria y Pintura";
                if (concepto61.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto61), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto62 = hoja.Celdas[82, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Materiales Diversos Hojalateria y Pintura.";
                if (concepto62.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto62), 0));
                }
                //_db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Materiales Diversos Hojalateria y Pintura.";
                if (concepto62.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto62), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto63 = hoja.Celdas[86, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 756;
                v1.Concepto = "Ingresos por Venta de Servicio";
                if (concepto63.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto63), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 756;
                v2.Concepto = "Ingresos por Venta de Servicio";
                if (concepto63.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto63), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto64 = hoja.Celdas[87, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 756;
                v1.Concepto = "Ingresos por Venta de Servicio";
                if (concepto64.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto64), 0));
                }
                //_db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 756;
                v2.Concepto = "Ingresos por Venta de Servicio";
                if (concepto64.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto64), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto65 = hoja.Celdas[90, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 757;
                v1.Concepto = "Ingresos por Venta de Refacciones";
                if (concepto65.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto65), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 757;
                v2.Concepto = "Ingresos por Venta de Refacciones";
                if (concepto65.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto65), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto66 = hoja.Celdas[91, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 758;
                v1.Concepto = "Venta a Taller";
                if (concepto66.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto66), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 758;
                v2.Concepto = "Venta a Taller";
                if (concepto66.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto66), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto67 = hoja.Celdas[92, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 759;
                v1.Concepto = "Venta a Mayoreo";
                if (concepto67.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto67), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 759;
                v2.Concepto = "Venta a Mayoreo";
                if (concepto67.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto67), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto68 = hoja.Celdas[93, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 759;
                v1.Concepto = "Venta a Mostrador";
                if (concepto68.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto68), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 759;
                v2.Concepto = "Venta a Mostrador";
                if (concepto68.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto68), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto69 = hoja.Celdas[94, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Refacciones";
                if (concepto68.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto68), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "Utilidad Bruta Refacciones";
                if (concepto68.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto68), 0));
                }

               // _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto70 = hoja.Celdas[97, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 761;
                v1.Concepto = "Ingresos por Hojalateria y Pintura";
                if (concepto70.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto70), 0));
                }
                _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 761;
                v2.Concepto = "Ingresos por Hojalateria y Pintura";
                if (concepto70.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto70), 0));
                }

                _db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }

            //--------------------------------------------------------------------------------------------

            var concepto71 = hoja.Celdas[98, columnaMes]; //
            if (version == "V1")
            {
                v1 = new ProcesoResultadoMensualExcel();
                v1.IdAgencia = idAgencia;
                v1.Anio = anio;
                v1.IdMes = mes;
                v1.IdReporte = 8;
                v1.IdConcepto = 0;
                v1.Concepto = "Utilidad Bruta Hojalateria y Pintura";
                if (concepto71.ToString() != "")
                {
                    v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto71), 0));
                }
               // _db.Insert(14066, 0, v1);
                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
            }
            else
            {
                v2 = new ProcesoResultadoMensualExtralibrosExcel();
                v2.IdAgencia = idAgencia;
                v2.Anio = anio;
                v2.IdMes = mes;
                v2.IdReporte = 8;
                v2.IdConcepto = 0;
                v2.Concepto = "IUtilidad Bruta Hojalateria y Pintura";
                if (concepto71.ToString() != "")
                {
                    v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto71), 0));
                }

                //_db.Insert(14066, 0, v2);

                Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
            }





            foreach (object celda in hoja.Celdas)
            {
                if (celda != null)
                    valor = celda.ToString();

                if ((c == (mes + 2)) && r >= 3) // Columna y row comparacion
                {
                    switch (r)
                    {
                        case 3:
                            if (version == "V1")
                            {
                                v1 = new ProcesoResultadoMensualExcel();
                                v1.IdAgencia = idAgencia;
                                v1.Anio = anio;
                                v1.IdMes = mes;
                                v1.IdReporte = 8;
                                v1.IdConcepto = 41;
                                v1.Concepto = "Efectivo e Inversiones";
                                v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                                //_db.Insert(14066, 0, v1);
                                Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                            }
                            break;
                        case 4:

                            break;

                        default:
                            break;
                    }
                }

                //if ((c == (mes + 2)) && r >= 3) // Columna y row comparacion
                //{
                //    if (r == 3) //Cuentas por cobrar unidades nuevas
                //    {
                //        if (version == "V1")
                //        {
                //            v1 = new ProcesoResultadoMensualExcel();
                //            v1.IdAgencia = idAgencia;
                //            v1.Anio = anio;
                //            v1.IdMes = mes;
                //            v1.IdReporte = 8;
                //            v1.IdConcepto = 41;
                //            v1.Concepto = "Efectivo e Inversiones";
                //            v1.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                //            //_db.Insert(14066, 0, v1);

                //            Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                //        }
                //        else
                //        {
                //            v2 = new ProcesoResultadoMensualExtralibrosExcel();
                //            v2.IdAgencia = idAgencia;
                //            v2.Anio = anio;
                //            v2.IdMes = mes;
                //            v2.IdReporte = 8;
                //            v2.IdConcepto = 41;
                //            v2.Concepto = "Efectivo e Inversiones";
                //            v2.Valor = celda == null ? 0 : Convert.ToInt32(Math.Round(Convert.ToDouble(valor == "" ? "0" : valor), 0));

                //           // _db.Insert(14066, 0, v2);

                //            Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                //        }
                //    }
                //}

                c++;

                if (r == hoja.LastRow)
                    r = 1;

                if (c == hoja.LastColumn + 2)
                {
                    c = 1;
                    r++;
                }

            }
        }
        
        public void LeePestanaAG(DVAExcel.HojaDeTrabajo hoja)
        {
            string valor = "";
            int r = 3;
            int c = 1;
            int celdaBase = 4;

            ProcesoResultadoMensualExcel v1;
            ProcesoResultadoMensualExtralibrosExcel v2;
            var celdaCoordenadas = hoja.Celdas[3, 3];
            var getValuesRow = hoja.GetStrRow(3);
            var testt = hoja.GetRowValues(3);
            var lastROw = hoja.LastRow;
            var lastColumn_ = hoja.LastColumn;

            var columnaMes = mes + 1;

            #region Alternativa
            var conceptos = new Dictionary<int, string>
                         {
                             { 668, "VENTAS AGENCIA MES" },
                             { 669, "VENTAS AGENCIA ACUMULADO" },
                             { 670, "VENTAS FIDEICOMISO MES" },
                             { 671, "VENTAS FIDEICOMISO ACUMULADO" },
                             { 672, "TOTAL VENTAS CON FIDEICOMISO ACUMULADO" },
                             { 673, "UNIDADES NUEVAS VENDIDAS AGENCIA MES" },
                             { 674, "UNIDADES NUEVAS VENDIDAS AUTOFIN MES" },
                             { 675, "UNIDADES NUEVAS VENDIDAS TAXIS BAM MES" },
                             { 676, "TOTAL UNIDADES NUEVAS VENDIDAS MES" },
                             { 677, "TOTAL UNIDADES NUEVAS VENDIDAS ACUMULADO" },
                             { 678, "PROMEDIO UTILIDAD BRUTA POR UNIDAD NUEVA VENDIDA ACUMULADO" },
                             { 679, "UTILIDAD BRUTA PROMEDIO ACUMULADO" },
                             { 680, "GASTO CORPORATIVO DEL MES" },
                             { 681, "UTILIDAD ANTES DE FIDEICOMISO MES" },
                             { 682, "UTILIDAD FIDEICOMISO MES" }, 
                             { 683, "PARTIDAS EXTRAORDINARIAS MES" },
                             { 684, "UTILIDAD FINAL ANTES DE IMPUESTOS MES" },
                             { 685, "GASTO CORPORATIVO ACUMULADO" },
                             { 686, "UTILIDAD ANTES DE FIDEICOMISO ACUMULADO" },
                             { 687, "UTILIDAD FIDEICOMISO ACUMULADA" },
                             { 688, "PARTIDAS EXTRAORDINARIAS ACUMULADO" },
                             { 689, "UTILIDAD FINAL ANTES DE IMPUESTOS ACUMULADA" },
                             { 690, "CAPITAL INVERTIDO" },

                             {691,   "RENDIMIENTO ANTES DE FIDEICOMISO MES" },
                             {692,   "RENDIMIENTO ANTES DE FIDEICOMISO ACUMULADO"},
                             {693,   "PRODUCTIVIDAD DEL MES"},
                             {694,   "PRODUCTIVIDAD DEL EJERCICIO"},

                             {695,   "AUTOS NUEVOS"},
                             {696,   "AUTOS SEMINUEVOS"},
                             {697,   "SERVICIO"},
                             {698,   "HOJALATERIA Y PINTURA"},
                             {699,   "REFACCIONES"},
                             {700,   "ADMINISTRACION"},
                             {701,   "VIGILANCIA Y ASEO"},
                             {702,   "AUTOS NUEVOS"},
                             {703,   "AUTOS SEMINUEVOS"},
                             {704,   "SERVICIO"},
                             {705,   "HOJALATERIA Y PINTURA"},
                             {706,   "REFACCIONES"},
                             {707,   "ADMINISTRACION"},
                             {708,   "VIGILANCIA Y ASEO"}
                         };

            for (int i = 0; i < conceptos.Count; i++)
            {
                var valorPorcentaje = conceptos.Keys.ElementAt(i);
                object concepto = new object { };
                if (valorPorcentaje == 691 || valorPorcentaje == 693 || valorPorcentaje == 702)
                {
                    celdaBase = celdaBase + 2;
                    concepto = hoja.Celdas[celdaBase + i, columnaMes]; // Incrementa posicion CELDA                 
                }
                else if (valorPorcentaje == 695)
                {
                    celdaBase = celdaBase + 3;
                    concepto = hoja.Celdas[celdaBase + i, columnaMes];
                }
                else
                {
                    concepto = hoja.Celdas[celdaBase + i, columnaMes];
                }
                 
                
                var valorConcepto = 0M;
                if (concepto != null)
                {
                    valorConcepto = Convert.ToDecimal(concepto);

                    if (valorPorcentaje == 679 || valorConcepto < 1)
                    {
                        if (Convert.ToDecimal(concepto) != 0)
                        {
                            concepto = Convert.ToDecimal(concepto) * 100;
                        }
                       
                    }
                }
               

                if (concepto == null)
                {
                    if (valorPorcentaje == 670)
                    {
                        concepto = 0;
                        goto esConceptoVacio;
                    }

                    continue;
                }
                esConceptoVacio:
                if (version == "V1")
                {
                     v1 = new ProcesoResultadoMensualExcel();
                    v1.IdAgencia = idAgencia;
                    v1.Anio = anio;
                    v1.IdMes = mes;
                    v1.IdReporte = 10;
                    if (conceptos.ContainsKey(v1.IdConcepto))
                    {
                        v1.IdConcepto = conceptos.Keys.ElementAt(i); 
                        v1.Concepto = conceptos[v1.IdConcepto];
                    }

                    if (concepto.ToString() != "")
                    {
                       v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto), 0));
                        int tomk = Convert.ToInt32(Math.Round(Convert.ToDecimal(concepto), 2));
                        v1.IdConcepto = conceptos.Keys.ElementAt(i); 
                        v1.Concepto = conceptos[v1.IdConcepto];
                    }

                    _db.Insert(14066, 0, v1);
                    Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                }
                else
                {
                     v2 = new ProcesoResultadoMensualExtralibrosExcel();
                    v2.IdAgencia = idAgencia;
                    v2.Anio = anio;
                    v2.IdMes = mes;
                    v2.IdReporte = 10;
                    if (conceptos.ContainsKey(v2.IdConcepto))
                    {
                        v2.IdConcepto = conceptos.Keys.ElementAt(i); // Utilizar la clave del diccionario
                        v2.Concepto = conceptos[v2.IdConcepto];
                    }

                    if (concepto.ToString() != "")
                    {
                        v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto), 0));
                        v2.IdConcepto = conceptos.Keys.ElementAt(i); // Utilizar la clave del diccionario
                        v2.Concepto = conceptos[v2.IdConcepto];
                    }

                  _db.Insert(14066, 0, v2);
                    Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                }
            }
            #endregion

  
        }

        public void LeePestanaSC(DVAExcel.HojaDeTrabajo hoja)
        {
            int celdaBase = 9;
            var testMes = mes;

            var mes_Celda = 0;

            ProcesoResultadoMensualExcel v1;
            ProcesoResultadoMensualExtralibrosExcel v2;
            var celdaCoordenadas = hoja.Celdas[3, 3];
            var getValuesRow = hoja.GetStrRow(3);
            var testt = hoja.GetRowValues(3);
            var lastROw = hoja.LastRow;
            var lastColumn_ = hoja.LastColumn;

            var columnaMes = mes + 1;

            if (mes == 1)
            {
                anio = anio - 1;
                mes = mes - 1;
            }
            else
            {
               
            }

            //mes = 7;
            switch (mes)
            {
                case 1:
                    mes_Celda = 9;
                    break;
                case 2:
                    mes_Celda = 28;
                    break;
                case 3:
                    mes_Celda = 47;
                    break;
                case 4:
                    mes_Celda = 66;
                    break;
                case 5:
                    mes_Celda = 85;
                    break;
                case 6:
                    mes_Celda = 104;
                    break;
                case 7:
                    mes_Celda = 123;
                    break;
                case 8:
                    mes_Celda = 142;
                    break;
                case 9:
                    mes_Celda = 161;
                    break;
                case 10:
                    mes_Celda = 180;
                    break;
                case 11:
                    mes_Celda = 199;
                    break;
                case 12:
                    mes_Celda = 218;
                    break;
                default:
                    break;
            }

            var columConcep = 4;



            var valorCeldaMes = mes_Celda;
            #region Alternativa
            var meses_ = new Dictionary<int, string>
            {
                {1, "Enero" },
                {2, "Febrero" },
                {3, "Marzo" },
                {4, "Abril" },
                {5, "Mayo" },
                {6, "Junio" },
                {7, "Julio" },
                {8, "Agosto" },
                {9, "Septiembre" },
                {10, "Octubre" },
                {11, "Noviembre" },
                {12, "Diciembre" }, 
            };

            var conceptos = new Dictionary<int, string>
                         {
                          {10560, "CARTERA ACTIVA - DOCUMENTOS_POR_COBRAR                   "},
                          {10561, "MONTO POR VENCER - DOCUMENTOS_POR_COBRAR                 "},
                          {10562, "VENCIDO DE 1 A 30 DIAS - DOCUMENTOS_POR_COBRAR           "},
                          {10563, "VENCIDO DE 31 A 60 DIAS - DOCUMENTOS_POR_COBRAR          "},
                          {10564, "VENCIDO DE 61 A 90 DIAS - DOCUMENTOS_POR_COBRAR          "},
                          {10565, "VENCIDO DE 91 A 120 DIAS - DOCUMENTOS_POR_COBRAR         "},
                          {10566, "VENCIDO MAS DE 120 DIAS - DOCUMENTOS_POR_COBRAR          "},

                          {10567, "CARTERA ACTIVA - AUTOS_TRADICIONAL                       "},
                          {10568, "MONTO POR VENCER - AUTOS_TRADICIONAL                     "},
                          {10569, "VENCIDO DE 1 A 30 DIAS - AUTOS_TRADICIONAL               "},
                          {10570, "VENCIDO DE 31 A 60 DIAS - AUTOS_TRADICIONAL              "},
                          {10571, "VENCIDO DE 61 A 90 DIAS - AUTOS_TRADICIONAL              "},
                          {10572, "VENCIDO DE 91 A 120 DIAS - AUTOS_TRADICIONAL             "},
                          {10573, "VENCIDO MAS DE 120 DIAS - AUTOS_TRADICIONAL              "},

                          {10574, "CARTERA ACTIVA - AUTOS_FIDEICOMISO_Y_AUTOFIN             "},
                          {10575, "MONTO POR VENCER - AUTOS_FIDEICOMISO_Y_AUTOFIN           "},
                          {10576, "VENCIDO DE 1 A 30 DIAS - AUTOS_FIDEICOMISO_Y_AUTOFIN     "},
                          {10577, "VENCIDO DE 31 A 60 DIAS - AUTOS_FIDEICOMISO_Y_AUTOFIN    "},
                          {10578, "VENCIDO DE 61 A 90 DIAS - AUTOS_FIDEICOMISO_Y_AUTOFIN    "},
                          {10579, "VENCIDO DE 91 A 120 DIAS - AUTOS_FIDEICOMISO_Y_AUTOFIN   "},
                          {10580, "VENCIDO MAS DE 120 DIAS - AUTOS_FIDEICOMISO_Y_AUTOFIN    "},

                          {10581, "CARTERA ACTIVA - SERVICIOS                               "},
                          {10582, "MONTO POR VENCER - SERVICIOS                             "},
                          {10583, "VENCIDO DE 1 A 30 DIAS - SERVICIOS                       "},
                          {10584, "VENCIDO DE 31 A 60 DIAS - SERVICIOS                      "},
                          {10585, "VENCIDO DE 61 A 90 DIAS - SERVICIOS                      "},
                          {10586, "VENCIDO DE 91 A 120 DIAS - SERVICIOS                     "},
                          {10587, "VENCIDO MAS DE 120 DIAS1 - SERVICIOS                     "},

                          {10588, "CARTERA ACTIVA - REFACCIONES                             "},
                          {10589, "MONTO POR VENCER - REFACCIONES                           "},
                          {10590, "VENCIDO DE 1 A 30 DIAS - REFACCIONES                     "},
                          {10591, "VENCIDO DE 31 A 60 DIAS - REFACCIONES                    "},
                          {10592, "VENCIDO DE 61 A 90 DIAS - REFACCIONES                    "},
                          {10593, "VENCIDO DE 91 A 120 DIAS - REFACCIONES                   "},
                          {10594, "VENCIDO MAS DE 120 DIAS - REFACCIONES                    "},

                          {10595, "CARTERA ACTIVA - PLANTA                                  "},
                          {10596, "MONTO POR VENCER - PLANTA                                "},
                          {10597, "VENCIDO DE 1 A 30 DIAS - PLANTA                          "},
                          {10598, "VENCIDO DE 31 A 60 DIAS - PLANTA                         "},
                          {10599, "VENCIDO DE 61 A 90 DIAS - PLANTA                         "},
                          {10600, "VENCIDO DE 91 A 120 DIAS - PLANTA                        "},
                          {10601, "VENCIDO MAS DE 120 DIAS - PLANTA                         "},

                          //{10602, "CARTERA ACTIVA - JURIDICO                                "},
                          //{10603, "MONTO POR VENCER - JURIDICO                              "},
                          //{10604, "VENCIDO DE 1 A 30 DIAS - JURIDICO                        "},
                          //{10605, "VENCIDO DE 31 A 60 DIAS - JURIDICO                       "},
                          //{10606, "VENCIDO DE 61 A 90 DIAS - JURIDICO                       "},
                          //{10607, "VENCIDO DE 91 A 120 DIAS - JURIDICO                      "},
                          //{10608, "VENCIDO MAS DE 120 DIAS - JURIDICO                       "},
                         };


            for (int i = 0; i < conceptos.Count; i++)
            {
                var valorPorcentaje = conceptos.Keys.ElementAt(i);
                if (valorPorcentaje == 10567 || valorPorcentaje == 10574 || valorPorcentaje == 10581 || valorPorcentaje == 10588 || valorPorcentaje == 10595)
                {
                    columConcep = columConcep + 2;
                    mes_Celda = valorCeldaMes;
                }

                var nombreConcepto = conceptos.Values.ElementAt(i);
                object concepto = new object { };
                if (valorPorcentaje == 691 || valorPorcentaje == 693 || valorPorcentaje == 702)
                {
                    mes_Celda = mes_Celda + 2;
                    concepto = hoja.Celdas[mes_Celda + i, columnaMes]; // Incrementa posicion CELDA                 
                }
                else if (valorPorcentaje == 695)
                {
                    mes_Celda = mes_Celda + 3;
                    concepto = hoja.Celdas[mes_Celda + i, columnaMes];
                }
                else
                {
                    concepto = hoja.Celdas[mes_Celda, columConcep];
                    mes_Celda++;
                    //columConcep++;
                }

              

       

                if (concepto == null)
                {
                    if (valorPorcentaje == 670)
                    {
                        concepto = 0;
                        goto esConceptoVacio;
                    }

                    continue;
                }
                esConceptoVacio:
                if (version == "V1")
                {
                    v1 = new ProcesoResultadoMensualExcel();
                    v1.IdAgencia = idAgencia;
                    v1.Anio = anio;
                    v1.IdMes = mes;
                    v1.IdReporte = 13;
                    if (conceptos.ContainsKey(v1.IdConcepto))
                    {
                        v1.IdConcepto = conceptos.Keys.ElementAt(i);
                        v1.Concepto = conceptos[v1.IdConcepto];
                    }

                    if (concepto.ToString() != "")
                    {
                        v1.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto), 0));
                        int tomk = Convert.ToInt32(Math.Round(Convert.ToDecimal(concepto), 2));
                        v1.IdConcepto = conceptos.Keys.ElementAt(i);
                        v1.Concepto = conceptos[v1.IdConcepto];
                    }

                    _db.Insert(14066, 0, v1);
                    Console.WriteLine("[CONCEPTO]: " + v1.IdConcepto + "_" + v1.Concepto + " = " + v1.Valor);
                }
                else //V2
                {
                    v2 = new ProcesoResultadoMensualExtralibrosExcel();
                    v2.IdAgencia = idAgencia;
                    v2.Anio = anio;
                    v2.IdMes = mes;
                    v2.IdReporte = 13;
                    if (conceptos.ContainsKey(v2.IdConcepto))
                    {
                        v2.IdConcepto = conceptos.Keys.ElementAt(i); // Utilizar la clave del diccionario
                        v2.Concepto = conceptos[v2.IdConcepto];
                    }

                    if (concepto.ToString() != "")
                    {
                        v2.Valor = Convert.ToInt32(Math.Round(Convert.ToDouble(concepto), 0));
                        v2.IdConcepto = conceptos.Keys.ElementAt(i); // Utilizar la clave del diccionario
                        v2.Concepto = conceptos[v2.IdConcepto];
                    }

                    _db.Insert(14066, 0, v2);
                    Console.WriteLine("[CONCEPTO]: " + v2.IdConcepto + "_" + v2.Concepto + " = " + v2.Valor);
                }
            }
            #endregion
        }


    }
}
