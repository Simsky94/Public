using DVAModelsReflection;
using DVAModelsReflection.Models.GRAL;
using DVAModelsReflection.Models.NOM;
using DVAModelsReflection.Models.PERS;
using DVAModelsReflectionFINA.Models.FINA;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DVAModelsReflection.Models.COMI.ConfiguracionPagadorasDomain;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    public class ProcesoGeneraComparacionExcelVsWeb
    {
        DB2Database _db = null;
        DVADB.DB2 dbCnx = new DVADB.DB2();

        int idAgencia = 0;
        int anio = 0;
        int mes = 0;
        string siglas = "";
        string query = "";


        string ruta = @"D:\Users\jasoria\Desktop\Escritorio\PMO\Proyectos 2023\ARCHIVO BASE\ARCHIVOS BASE EXCEL\";
        //string ruta = @"C:\Users\fevangelista\Documents\Reportes Cuaderno Contable\";

        public ProcesoGeneraComparacionExcelVsWeb(DB2Database _db)
        {
            this._db = _db;
        }

        public ProcesoGeneraComparacionExcelVsWeb(DB2Database _db, int aAnio, int aMes, string aSiglas)
        {
            this._db = _db;
            this.mes = aMes;
            this.anio = aAnio;
            this.siglas = aSiglas;

            ruta = ruta + anio + "\\" + mes + "\\" + "ARCHIVO BASE " + siglas + "_COMPARACIÓN EXCEL VS WEB_" + anio + "_" + mes + ".xlsx";

            Console.WriteLine("[INICIA GENERACION EXCEL COMPARACIÓN]: ");
            Console.WriteLine("[AÑO]: " + anio);
            Console.WriteLine("[MES]: " + mes);
        }

        public ProcesoGeneraComparacionExcelVsWeb(DB2Database _db, int aAnio, int aMes, string aSiglas, List<int> idsAgencia)
        {
            this._db = _db;
            this.mes = aMes;
            this.anio = aAnio;
            this.siglas = aSiglas;

            List<AgenciasReportes> agenciasReportes = AgenciasReportes.ListarPorIds(_db, idsAgencia);
            List<Agencia> agencias = Agencia.ListarPorIds(_db, idsAgencia);

            ruta = ruta + anio + "\\" + mes + "\\" + "ARCHIVO BASE " + siglas + "_COMPARACIÓN EXCEL VS WEB_" + anio + "_" + mes + "_";

            foreach (Agencia age in agencias)
            {
                ruta = ruta + age.Siglas + "_";
            }

            ruta += ".xlsx";

            Console.WriteLine("[INICIA GENERACION EXCEL COMPARACIÓN]: ");
            Console.WriteLine("[AÑO]: " + anio);
            Console.WriteLine("[MES]: " + mes);
        }

        public void GeneraExcelRM(List<int> idsAgencia)
        {
            List<Agencia> agencias;

            if (idsAgencia == null)
            {
                List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1);
                List<int> aIdAgencias = new List<int>();
                aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                agencias = Agencia.ListarPorIds(_db, aIdAgencias);
            }
            else
                agencias = Agencia.ListarPorIds(_db, idsAgencia);

            List<ConceptosContables> conceptosV1 = ConceptosContables.ListarRMV1(_db);
            List<ConceptosContables> conceptosV2 = ConceptosContables.ListarRMV2(_db);
            DataTable dtRMExcelV1 = GetExcelV1(5);
            DataTable dtRMExcelV1Acum = GetExcelV1Acumulado(5);
            DataTable dtRMExcelV2 = GetExcelV2(5);
            DataTable dtRMExcelV2Acum = GetExcelV2Acumulado(5);
            DataTable dtRMWebV1 = GetRMWebV1();
            DataTable dtRMWebV1Acum = GetRMWebV1Acumulado();
            DataTable dtRMWebV2 = GetRMWebV2();
            DataTable dtRMWebV2Acum = GetRMWebV2Acumulado();

            int r = 1;
            int c = 1;
            int vExcel = 0;
            int vWeb = 0;
            int intAux = 0;
            DataRow[] drRMExcelV1 = null;
            DataRow[] drRMExcelV1Acum = null;
            DataRow[] drRMExcelV2 = null;
            DataRow[] drRMExcelV2Acum = null;
            DataRow[] drRMWebV1 = null;
            DataRow[] drRMWebV1Acum = null;
            DataRow[] drRMWebV2 = null;
            DataRow[] drRMWebV2Acum = null;

            List<Reporte> listaR = new List<Reporte>();
            List<ReporteGlobal> listaRG = new List<ReporteGlobal>();
            ReporteGlobal rG;

            using (DVAExcel.ExcelWriter eW = new DVAExcel.ExcelWriter(ruta))
            {
                foreach (Agencia agencia in agencias.OrderBy(o => o.Siglas))
                {
                    if ((agencia.Id == 590) || (agencia.Id == 583) || (agencia.Id == 301) || (agencia.Id == 583) || (agencia.Id == 563) || (agencia.Id == 593)
                        || (agencia.Id == 592) || (agencia.Id == 594) || (agencia.Id == 100))
                        continue;

                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("[SIGLAS]: " + agencia.Siglas);
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");

                    foreach (ConceptosContables concepto in conceptosV2)
                    {
                        Reporte rep = new Reporte();

                        r++;

                        rep.ID_CONCEPTO = concepto.Id;
                        rep.CONCEPTO = concepto.NombreConcepto;

                        drRMExcelV1 = dtRMExcelV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drRMExcelV1Acum = dtRMExcelV1Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        if (agencia.Id == 27)
                        {
                            drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 286)
                        {
                            drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",100) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",100) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 36)
                        {
                            drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 12)
                        {
                            drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 35)
                        {
                            drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",595) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",595) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 588)
                        {
                            drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",590) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",590) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 32)
                        {
                            drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",116) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",116) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 212)
                        {
                            drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",33) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",33) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else
                        {
                            drRMWebV1 = dtRMWebV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV1Acum = dtRMWebV1Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        }

                        drRMExcelV2 = dtRMExcelV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drRMExcelV2Acum = dtRMExcelV2Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        if (agencia.Id == 27)
                        {
                            drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 286)
                        {
                            drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",100) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",100) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 36)
                        {
                            drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 12)
                        {
                            drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 35)
                        {
                            drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",595) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",595) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 588)
                        {
                            drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",590) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",590) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 32)
                        {
                            drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",116) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",116) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 212)
                        {
                            drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",33) AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA IN (" + agencia.Id + ",33) AND ID_CONCEPTO = " + concepto.Id);
                        }
                        else
                        {
                            drRMWebV2 = dtRMWebV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                            drRMWebV2Acum = dtRMWebV2Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        }

                        if (drRMExcelV1.Length != 0)
                        {
                            vExcel = Convert.ToInt32(drRMExcelV1[0]["VALOR"]);
                            rep.EXCEL_V1 = vExcel;
                            vExcel = Convert.ToInt32(drRMExcelV1Acum[0]["VALOR"]);
                            rep.EXCEL_V1_ACUM = vExcel;
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
                                    FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, agencia.Id, FINA_EParametros.HorasFacturadasServVal);

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
                                    FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, agencia.Id, FINA_EParametros.HorasFacturadasServVal);

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

                        if (drRMExcelV2.Length != 0)
                        {
                            vExcel = Convert.ToInt32(drRMExcelV2[0]["VALOR"]);
                            rep.EXCEL_V2 = vExcel;
                            vExcel = Convert.ToInt32(drRMExcelV2Acum[0]["VALOR"]);
                            rep.EXCEL_V2_ACUM = vExcel;
                        }

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
                                    FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, agencia.Id, FINA_EParametros.HorasFacturadasServVal);

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
                                    FINA_AgenciaParametro resp1 = FINA_AgenciaParametro.Buscar(_db, agencia.Id, FINA_EParametros.HorasFacturadasServVal);

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

                        Console.WriteLine("[CONCEPTO]: " + concepto.Id + "_" + concepto.NombreConcepto +
                            " [EXCEL_V1]: " + rep.EXCEL_V1 +
                            " [WEB_V1]: " + rep.WEB_V1 +
                            " [DIFF_V1]: " + rep.DIFF_V1 +
                            " [EXCEL_V1_ACUMULADO]: " + rep.EXCEL_V1_ACUM +
                            " [WEB_V1_ACUMULADO]: " + rep.WEB_V1_ACUM +
                            " [DIFF_V1_ACUMULADO]: " + rep.DIFF_V1_ACUM +
                            " [EXCEL_V2]: " + rep.EXCEL_V2 +
                            " [WEB_V2]: " + rep.WEB_V2 +
                            " [DIFF_V2]: " + rep.DIFF_V2 +
                            " [EXCEL_V2_ACUMULADO]: " + rep.EXCEL_V2_ACUM +
                            " [WEB_V2_ACUMULADO]: " + rep.WEB_V2_ACUM +
                            " [DIFF_V2_ACUMULADO]: " + rep.DIFF_V2_ACUM);

                        listaR.Add(rep);
                    }

                    eW.WriteDataTable(listaR.ToDataTable(), agencia.Siglas);
                    //eW.Dispose();

                    listaR = new List<Reporte>();
                }

                //foreach (ConceptosContables concepto in conceptosV2)
                //{
                //    drRMExcelV1 = dtRMExcelV1.Select("ID_CONCEPTO = " + concepto.Id);
                //    drRMWebV1 = dtRMWebV1.Select("ID_CONCEPTO = " + concepto.Id);

                //    drRMExcelV2 = dtRMExcelV2.Select("ID_CONCEPTO = " + concepto.Id);
                //    drRMWebV2 = dtRMWebV2.Select("ID_CONCEPTO = " + concepto.Id);

                //    foreach (DataRow dr in drRMExcelV1)
                //    {
                //        if (agencia.Siglas == "AAZ")
                //        {
                //            rG = new ReporteGlobal();

                //            rG.AAZ_ID_CONCEPTO = concepto.Id;
                //            rG.AAZ_CONCEPTO = concepto.NombreConcepto;
                //            rG.AAZ_EXCEL_V1 = rep.EXCEL_V1;
                //            rG.AAZ_WEB_V1 = rep.WEB_V1;
                //            rG.AAZ_DIFF_V1 = rep.DIFF_V1;
                //            rG.AAZ_EXCEL_V2 = rep.EXCEL_V2;
                //            rG.AAZ_WEB_V2 = rep.WEB_V2;
                //            rG.AAZ_DIFF_V2 = rep.DIFF_V2;

                //            listaRG.Add(rG);
                //        }
                //        else if (agencia.Siglas == "AEP")
                //        {
                //            rG = new ReporteGlobal();

                //            rG.AEP_ID_CONCEPTO = concepto.Id;
                //            rG.AEP_CONCEPTO = concepto.NombreConcepto;
                //            rG.AEP_EXCEL_V1 = rep.EXCEL_V1;
                //            rG.AEP_WEB_V1 = rep.WEB_V1;
                //            rG.AEP_DIFF_V1 = rep.DIFF_V1;
                //            rG.AEP_EXCEL_V2 = rep.EXCEL_V2;
                //            rG.AEP_WEB_V2 = rep.WEB_V2;
                //            rG.AEP_DIFF_V2 = rep.DIFF_V2;

                //            listaRG.Add(rG);
                //        }
                //        else if (agencia.Siglas == "ASP")
                //        {
                //            rG = new ReporteGlobal();

                //            rG.ASP_ID_CONCEPTO = concepto.Id;
                //            rG.ASP_CONCEPTO = concepto.NombreConcepto;
                //            rG.ASP_EXCEL_V1 = rep.EXCEL_V1;
                //            rG.ASP_WEB_V1 = rep.WEB_V1;
                //            rG.ASP_DIFF_V1 = rep.DIFF_V1;
                //            rG.ASP_EXCEL_V2 = rep.EXCEL_V2;
                //            rG.ASP_WEB_V2 = rep.WEB_V2;
                //            rG.ASP_DIFF_V2 = rep.DIFF_V2;

                //            listaRG.Add(rG);
                //        }
                //    }
                //}

                //eW.WriteDataTable(listaRG.ToDataTable(), "GLOBAL");
                //eW.Dispose();

                eW.Dispose();
            }
        }

        public DataTable GetExcelV1(int idReporte)
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RESMEX . FIFNYEAR ANIO, RESMEX . FIFNMONTH MES, RESMEX . FIFNVALUE VALOR, CONCP . FIFNAGRPM ORDEN\r\n" +
                "FROM PRODFINA.FNDRESMEX RESMEX\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RESMEX . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RESMEX . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RESMEX . FIFNYEAR = " + anio + " AND RESMEX . FIFNMONTH = " + mes + " " +
                "AND RESMEX.FIFNIDRPT = " + idReporte;

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetExcelV1Acumulado(int idReporte)
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RESMEX . FIFNYEAR ANIO, SUM(RESMEX . FIFNVALUE) VALOR, CONCP . FIFNAGRPM ORDEN\r\n" +
                "FROM PRODFINA.FNDRESMEX RESMEX\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RESMEX . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RESMEX . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RESMEX . FIFNYEAR = " + anio + " AND RESMEX . FIFNMONTH <= " + mes + " " +
                "AND RESMEX.FIFNIDRPT = " + idReporte + "\r\n" +
                "GROUP BY CIAUN . FIGEIDCIAU, CIAUN . FSGERAZSOC, CIAUN . FSGESIGCIA, CONCP . FIFNIDCPT, CONCP . FSFNNCPT, RESMEX . FIFNYEAR, CONCP . FIFNAGRPM";

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetExcelV2(int idReporte)
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RSMEEX . FIFNYEAR ANIO, RSMEEX . FIFNMONTH MES, RSMEEX . FIFNVALUE VALOR, CONCP . FIFNORDEV2 ORDEN\r\n" +
                "FROM PRODFINA.FNDRSMEEX RSMEEX\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RSMEEX . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RSMEEX . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RSMEEX . FIFNYEAR = " + anio + " AND RSMEEX . FIFNMONTH = " + mes + " " +
                "AND RSMEEX.FIFNIDRPT = " + idReporte;

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetExcelV2Acumulado(int idReporte)
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RSMEEX . FIFNYEAR ANIO, SUM(RSMEEX . FIFNVALUE) VALOR, CONCP . FIFNORDEV2 ORDEN\r\n" +
                "FROM PRODFINA.FNDRSMEEX RSMEEX\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RSMEEX . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RSMEEX . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RSMEEX . FIFNYEAR = " + anio + " AND RSMEEX . FIFNMONTH <= " + mes + "\r\n" +
                "AND RSMEEX.FIFNIDRPT = " + idReporte + "\r\n" +
                "GROUP BY CIAUN . FIGEIDCIAU, CIAUN . FSGERAZSOC, CIAUN . FSGESIGCIA, CONCP . FIFNIDCPT, CONCP . FSFNNCPT, RSMEEX . FIFNYEAR, CONCP . FIFNORDEV2";

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetRMWebV1()
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RESMEN . FIFNYEAR ANIO, RESMEN . FIFNMONTH MES, RESMEN . FIFNVALUE VALOR, CONCP . FIFNAGRPM ORDEN\r\n" +
                "FROM PRODFINA.FNDRESMEN RESMEN\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RESMEN . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RESMEN . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RESMEN . FIFNYEAR = " + anio + " AND RESMEN . FIFNMONTH = " + mes;

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetRMWebV1Acumulado()
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RESMEN . FIFNYEAR ANIO, SUM(RESMEN . FIFNVALUE) VALOR, CONCP . FIFNAGRPM ORDEN\r\n" +
                "FROM PRODFINA.FNDRESMEN RESMEN\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RESMEN . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RESMEN . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RESMEN . FIFNYEAR = " + anio + " AND RESMEN . FIFNMONTH <= " + mes + "\r\n" +
                "GROUP BY CIAUN . FIGEIDCIAU, CIAUN . FSGERAZSOC, CIAUN . FSGESIGCIA, CONCP . FIFNIDCPT, CONCP . FSFNNCPT, RESMEN . FIFNYEAR, CONCP . FIFNAGRPM";

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetRMWebV2()
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RSMENE . FIFNYEAR ANIO, RSMENE . FIFNMONTH MES, RSMENE . FIFNVALUE VALOR, CONCP . FIFNORDEV2 ORDEN\r\n" +
                "FROM PRODFINA.FNDRSMENE RSMENE\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RSMENE . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RSMENE . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RSMENE . FIFNYEAR = " + anio + " AND RSMENE . FIFNMONTH = " + mes;

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetRMWebV2Acumulado()
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RSMENE . FIFNYEAR ANIO, SUM(RSMENE . FIFNVALUE) VALOR, CONCP . FIFNORDEV2 ORDEN\r\n" +
                "FROM PRODFINA.FNDRSMENE RSMENE\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RSMENE . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RSMENE . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RSMENE . FIFNYEAR = " + anio + " AND RSMENE . FIFNMONTH <= " + mes + "\r\n" +
                "GROUP BY CIAUN . FIGEIDCIAU, CIAUN . FSGERAZSOC, CIAUN . FSGESIGCIA, CONCP . FIFNIDCPT, CONCP . FSFNNCPT, RSMENE . FIFNYEAR, CONCP . FIFNORDEV2";

            return dbCnx.GetDataTable(query);
        }

        public void GeneraExcelBG(List<int> idsAgencia)
        {
            List<Agencia> agencias;

            if (idsAgencia == null)
            {
                List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1); //Todas las agencias
                List<int> aIdAgencias = new List<int>();
                aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                agencias = Agencia.ListarPorIds(_db, aIdAgencias);
            }
            else
                agencias = Agencia.ListarPorIds(_db, idsAgencia);

            List<ConceptosContables> conceptosV1 = ConceptosContables.ListarBGV1yV2(_db); //obtiene el listado de todos los conceptos contables
            List<ConceptosContables> conceptosV2 = ConceptosContables.ListarBGV1yV2(_db);

            DataTable dtBGExcelV1 = GetExcelV1(4);
            DataTable dtBGExcelV1Acum = GetExcelV1Acumulado(4);
            DataTable dtBGExcelV2 = GetExcelV2(4);
            DataTable dtBGExcelV2Acum = GetExcelV2Acumulado(4);
            DataTable dtBGWebV1 = GetBGWebV1();
            //DataTable dtBGWebV1Acum = GetBGWebV1();
            DataTable dtBGWebV2 = GetBGWebV2();
            //DataTable dtBGWebV2Acum = GetBGWebV1();

            int r = 1;
            int c = 1;
            int vExcel = 0;
            int vWeb = 0;
            int intAux = 0;
            DataRow[] drBGExcelV1 = null;
            DataRow[] drBGExcelV1Acum = null;
            DataRow[] drBGExcelV2 = null;
            DataRow[] drBGExcelV2Acum = null;
            DataRow[] drBGWebV1 = null;
            DataRow[] drBGWebV1Acum = null;
            DataRow[] drBGWebV2 = null;
            DataRow[] drBGWebV2Acum = null;

            List<Reporte> listaR = new List<Reporte>();
            List<ReporteGlobal> listaRG = new List<ReporteGlobal>();
            ReporteGlobal rG;

            using (DVAExcel.ExcelWriter eW = new DVAExcel.ExcelWriter(ruta))
            {
                foreach (Agencia agencia in agencias.OrderBy(o => o.Siglas))
                {
                    try
                    {
                        if ((agencia.Id == 590) || (agencia.Id == 583) || (agencia.Id == 301) || (agencia.Id == 583) || (agencia.Id == 563) || (agencia.Id == 593)
                            || (agencia.Id == 592) || (agencia.Id == 594) || (agencia.Id == 100))
                            continue;

                        Console.WriteLine("******************************************************************************************************");
                        Console.WriteLine("******************************************************************************************************");
                        Console.WriteLine("******************************************************************************************************");
                        Console.WriteLine("[SIGLAS]: " + agencia.Siglas);
                        Console.WriteLine("******************************************************************************************************");
                        Console.WriteLine("******************************************************************************************************");
                        Console.WriteLine("******************************************************************************************************");

                        foreach (ConceptosContables concepto in conceptosV2)
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

                            drBGExcelV1 = dtBGExcelV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                            drBGExcelV1Acum = dtBGExcelV1Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                            if (agencia.Id == 27)
                            {
                                drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 286)
                            {
                                drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",100) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",100) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 36)
                            {
                                drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 12)
                            {
                                drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 35)
                            {
                                drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",595) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",595) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 588)
                            {
                                drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",590) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",590) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 32)
                            {
                                drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",116) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",116) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 212)
                            {
                                drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",33) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA IN (" + agencia.Id + ",33) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else
                            {
                                drBGWebV1 = dtBGWebV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV1Acum = dtBGWebV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                            }

                            drBGExcelV2 = dtBGExcelV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                            drBGExcelV2Acum = dtBGExcelV2Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                            if (agencia.Id == 27)
                            {
                                drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",301,583) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 286)
                            {
                                drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",100) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",100) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 36)
                            {
                                drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",563,593) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 12)
                            {
                                drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",592,594) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 35)
                            {
                                drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",595) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",595) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 588)
                            {
                                drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",590) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",590) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 32)
                            {
                                drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",116) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",116) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else if (agencia.Id == 212)
                            {
                                drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",33) AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA IN (" + agencia.Id + ",33) AND ID_CONCEPTO = " + concepto.Id);
                            }
                            else
                            {
                                drBGWebV2 = dtBGWebV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                                drBGWebV2Acum = dtBGWebV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                            }

                            if (drBGExcelV1.Length != 0)
                            {
                                vExcel = Convert.ToInt32(drBGExcelV1[0]["VALOR"]);
                                rep.EXCEL_V1 = vExcel;
                                vExcel = Convert.ToInt32(drBGExcelV1Acum[0]["VALOR"]);
                                rep.EXCEL_V1_ACUM = vExcel;
                            }

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

                            rep.DIFF_V1 = rep.EXCEL_V1 - rep.WEB_V1;
                            rep.DIFF_V1_ACUM = rep.EXCEL_V1_ACUM - rep.WEB_V1_ACUM;

                            if (drBGExcelV2.Length != 0)
                            {
                                vExcel = Convert.ToInt32(drBGExcelV2[0]["VALOR"]);
                                rep.EXCEL_V2 = vExcel;
                                vExcel = Convert.ToInt32(drBGExcelV2Acum[0]["VALOR"]);
                                rep.EXCEL_V2_ACUM = vExcel;
                            }

                            if (drBGWebV2.Length != 0)
                            {
                                vWeb = 0;

                                foreach (DataRow dr in drBGWebV2)
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

                                    rep.WEB_V2 = vWeb * cambioSigo;
                                }

                                vWeb = 0;

                                foreach (DataRow dr in drBGWebV2Acum)
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

                                    rep.WEB_V2_ACUM = vWeb * cambioSigo;
                                }
                            }

                            rep.DIFF_V2 = rep.EXCEL_V2 - rep.WEB_V2;
                            rep.DIFF_V2_ACUM = rep.EXCEL_V2_ACUM - rep.WEB_V2_ACUM;

                            Console.WriteLine("[CONCEPTO]: " + concepto.Id + "_" + concepto.NombreConcepto +
                                " [EXCEL_V1]: " + rep.EXCEL_V1 +
                                " [WEB_V1]: " + rep.WEB_V1 +
                                " [DIFF_V1]: " + rep.DIFF_V1 +
                                " [EXCEL_V1_ACUMULADO]: " + rep.EXCEL_V1_ACUM +
                                " [WEB_V1_ACUMULADO]: " + rep.WEB_V1_ACUM +
                                " [DIFF_V1_ACUMULADO]: " + rep.DIFF_V1_ACUM +
                                " [EXCEL_V2]: " + rep.EXCEL_V2 +
                                " [WEB_V2]: " + rep.WEB_V2 +
                                " [DIFF_V2]: " + rep.DIFF_V2 +
                                " [EXCEL_V2_ACUMULADO]: " + rep.EXCEL_V2_ACUM +
                                " [WEB_V2_ACUMULADO]: " + rep.WEB_V2_ACUM +
                                " [DIFF_V2_ACUMULADO]: " + rep.DIFF_V2_ACUM);

                            listaR.Add(rep);
                        }

                        eW.WriteDataTable(listaR.ToDataTable(), agencia.Siglas);                        

                        listaR = new List<Reporte>();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        Console.WriteLine(ex.InnerException);
                        Console.WriteLine(ex.StackTrace);
                    }
                }

                eW.Dispose();
            }
        }

        public DataTable GetBGWebV1()
        {
            query = "select \r\n                            ID_AGENCIA,\r\n                            IDGRUPO\r\n                            ,NOMBREGRUPOGENERAL\r\n\r\n                            ,CASE WHEN IDGRUPOMASTER = 2 THEN 1\r\n                                  WHEN  IDGRUPOMASTER = 1365 THEN 2\r\n                                  WHEN  IDGRUPOMASTER = 3 THEN 3\r\n                                  WHEN  IDGRUPOMASTER = 14 THEN 4\r\n                                  ELSE IDGRUPOMASTER END IDGRUPOMASTER\r\n                                  \r\n                            ,NOMBREGRUPOMAESTRO\r\n                            \r\n                            ,INDEXGRUPO\r\n                            ,IDGRIPOGEN\r\n                            ,NOMBREGRUPOAGRUPADOR\r\n                            ,ID_CONCEPTO\r\n                            ,NOMBRE_CONCEPTO\r\n\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 1 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 1 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_ENERO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 2 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 2 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_FEBRERO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 3 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 3 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_MARZO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 4 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 4 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_ABRIL\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 5 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 5 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_MAYO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 6 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 6 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_JUNIO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 7 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 7 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_JULIO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 8 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 8 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_AGOSTO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 9 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 9 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_SEPTIEMBRE\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 10 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 10 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_OCTUBRE\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 11 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 11 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_NOVIEMBRE\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 12 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 12 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_DICIEMBRE\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(coalesce(Saldo_CUENTA,0) + coalesce(Saldo_CARTERA,0))*-1 ELSE sum(coalesce(Saldo_CUENTA,0) + coalesce(Saldo_CARTERA,0)) END TOTAL\r\n\r\n                            FROM (\r\n\r\n                            SELECT  distinct\r\n                            TB1.FICOMES\r\n                            ,TB1.FIFNIDCIAU ID_AGENCIA\r\n                            ,(select TRIM(FSGERAZSOC)  FROM PRODGRAL.GECCIAUN WHERE FIGEIDCIAU = TB1.FIFNIDCIAU AND FIGESTATUS = 1 ) RAZON_SOCIAL\r\n                            ,(select TRIM(FSGESIGCIA)  FROM PRODGRAL.GECCIAUN WHERE FIGEIDCIAU = TB1.FIFNIDCIAU AND FIGESTATUS = 1 ) FSGESIGCIA\r\n                            ,TRIM(FSFNRPTNM) REPORTE\r\n                            ,FIFNIDAGRP   IDGRUPO\r\n                            ,TRIM(FSFNIDAGRP) NOMBREGRUPOGENERAL \r\n                            ,FIFNIDGRPM  IDGRUPOMASTER\r\n                            ,TRIM(FSFNAGRPM)  NOMBREGRUPOMAESTRO\r\n                            ,INDEXGRUPO \r\n                            ,GRUPO  IDGRIPOGEN   \r\n                            ,TRIM(NOMBREGRUPO) NOMBREGRUPOAGRUPADOR\r\n                            ,TB1.ID_CONCEPTO\r\n                            ,TRIM(TB1.CONCEPTO) NOMBRE_CONCEPTO\r\n                            ,TB1.IDCUENTA\r\n                            ,SUBSTR(TB1.DESC_CTA,1,LOCATE(' ', TB1.DESC_CTA)) Key_CTA\r\n                            ,trim(TB1.DESC_CTA) DESCRIPCION_CUENTA\r\n                            ,(ifnull(CTAS.FDCOSALDOF,0)) Saldo_CUENTA\r\n                            ,CARTERAS.FICCIDCART\r\n                            ,TRIM(CARDESC.FSCCCVECAR) CLAVE_CARTERA\r\n                            ,TRIM(CARDESC.FSCCDESCAR) DESCRIPCION_CARTERA\r\n                            ,CASE WHEN CARTERAS.FICCIDCART = 0 THEN 0 ELSE COALESCE(CARTERAS.FDCCSALDOF,0) END Saldo_CARTERA\r\n                            FROM (\r\n                            SELECT DISTINCT \r\n\r\n                            0 FIFNIDDEP,\r\n                            PERI.FDCOIDPERI,\r\n                            PERI.FICOMES,\r\n                            PERI.FICOANIO,\r\n                            CONFIGREP.FIFNIDCIAU,\r\n                            CONFIGREP.FIFNIDRPT,\r\n                            REPORTES.FSFNRPTNM,\r\n                            CONFIGREP.FIFNIDAGRP,\r\n                            AGRPGENERAL.FSFNIDAGRP,\r\n                            CONFIGREP.FIFNIDGRPM,\r\n                            AGRPMAESTRO.FSFNAGRPM,\r\n                            AGRUPADOR.FIFNINDEX AS INDEXGRUPO,\r\n                            CONFIGREP.FIFNIDGRP AS GRUPO,\r\n                            AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,\r\n                            CONFIGREP.FIFNIDCPT AS ID_CONCEPTO,\r\n                            CONCEPTOS.FSFNNCPT AS CONCEPTO,\r\n                            COALESCE(CONFIG.FIFNIDCART,0) IDCART,\r\n                            COALESCE(CONFIG.FIFNIDCTA,0) IDCUENTA,\r\n                            (SELECT trim(CUENTA.FSCOCUENTA) || ' ' || CUENTA.FSCODESCTA FROM PRODCONT.COCATCTS CUENTA WHERE CUENTA.FICOIDTCCT = 1 AND cuenta.FICOIDCTA = CONFIG.FIFNIDCTA FETCH FIRST 1 ROW ONLY) DESC_CTA\r\n\r\n                            FROM \r\n                            PRODCONT.COCPERIO PERI, --PRODCONT.COCPERIO periodos\r\n                            PRODFINA.FNDCFREPCP CONFIGREP --PRODFINA.FNDCFREPCP Configuración del reporte cpt sus agrupadores\r\n                            LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP --PRODFINA.FNCAGRGRAL Agrupador general\r\n                            LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM --PRODFINA.FNCAGRUPA Agrupador maestro\r\n                            LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP --PRODFINA.FNCGRPCT Grupo o AGRUPADOR\r\n                            LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT AND CONCEPTOS.FIFNSTATUS = 1 --PRODFINA.FNCCONCP Conceptos\r\n                            LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT --PRODFINA.FNCREPORT Catalogo de reportes\r\n                            LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNIDTCCT IN (1) AND CONFIG.FIFNSTATUS = 1\r\n                            WHERE \r\n                            PERI.FICOANIO = " + anio + "\r\n                            --AND PERI.FICOMES = 12\r\n                            and REPORTES.FIFNIDRPT IN (4)\r\n--                            AND CONFIGREP.FIFNIDCIAU in (28) \r\n                            AND CONCEPTOS.FIFNSTATUS = 1\r\n                            AND REPORTES.FIFNSTATUS = 1\r\n                            AND CONFIG.FIFNSTATUS = 1\r\n                            ORDER BY \r\n                            CONFIGREP.FIFNIDRPT,\r\n                            CONFIGREP.FIFNIDCPT \r\n                            ) TB1 \r\n                            LEFT JOIN PRODCXC.CCESALCA CARTERAS ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU \r\n                            AND CARTERAS.FICCIDCART = TB1.IDCART \r\n                            AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI\r\n                            AND CARTERAS.FICCSTATUS = 1\r\n                            LEFT join PRODCXC.CCCARTER CARDESC on CARTERAS.FICCIDCART = CARDESC.FICCIDCART AND CARDESC.FICCSTATUS = 1\r\n\r\n\r\n                            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI \r\n                            AND CTAS.FICOIDTCCT IN (1) \r\n                            AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU \r\n                            AND CTAS.FICOIDCTA = TB1.IDCUENTA \r\n                            AND CTAS.FICOSTATUS = 1 \r\n\r\n\r\n                            ) T1\r\n\r\n\r\n                            group BY\r\n                            ID_AGENCIA\r\n\r\n                            ,IDGRUPO\r\n                            ,NOMBREGRUPOGENERAL\r\n\r\n                            ,IDGRUPOMASTER\r\n                            ,NOMBREGRUPOMAESTRO\r\n\r\n                            ,INDEXGRUPO\r\n                            ,IDGRIPOGEN\r\n                            ,NOMBREGRUPOAGRUPADOR\r\n                            ,ID_CONCEPTO\r\n                            ,NOMBRE_CONCEPTO\r\n\r\n                            ORDER BY \r\n                            T1.IDGRUPO ASC";

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetBGWebV2()
        {
            query = "select \r\n                            ID_AGENCIA,\r\n                            IDGRUPO\r\n                            ,NOMBREGRUPOGENERAL\r\n\r\n                            ,CASE WHEN IDGRUPOMASTER = 2 THEN 1\r\n                                  WHEN  IDGRUPOMASTER = 1365 THEN 2\r\n                                  WHEN  IDGRUPOMASTER = 3 THEN 3\r\n                                  WHEN  IDGRUPOMASTER = 14 THEN 4\r\n                                  ELSE IDGRUPOMASTER END IDGRUPOMASTER\r\n                                  \r\n                            ,NOMBREGRUPOMAESTRO\r\n                            \r\n                            ,INDEXGRUPO\r\n                            ,IDGRIPOGEN\r\n                            ,NOMBREGRUPOAGRUPADOR\r\n                            ,ID_CONCEPTO\r\n                            ,NOMBRE_CONCEPTO\r\n\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 1 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 1 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_ENERO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 2 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 2 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_FEBRERO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 3 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 3 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_MARZO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 4 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 4 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_ABRIL\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 5 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 5 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_MAYO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 6 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 6 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_JUNIO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 7 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 7 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_JULIO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 8 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 8 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_AGOSTO\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 9 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 9 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_SEPTIEMBRE\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 10 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 10 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_OCTUBRE\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 11 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 11 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_NOVIEMBRE\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(CASE WHEN FICOMES = 12 THEN  COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0) ELSE 0 END )*-1 ELSE sum(CASE WHEN FICOMES = 12 THEN  (COALESCE(Saldo_CUENTA,0) + COALESCE(Saldo_CARTERA,0)) / 1000 ELSE 0 END ) END SALDO_DICIEMBRE\r\n                            ,CASE WHEN ID_CONCEPTO IN (10610) THEN sum(coalesce(Saldo_CUENTA,0) + coalesce(Saldo_CARTERA,0))*-1 ELSE sum(coalesce(Saldo_CUENTA,0) + coalesce(Saldo_CARTERA,0)) END TOTAL\r\n\r\n                            FROM (\r\n\r\n                            SELECT  distinct\r\n                            TB1.FICOMES\r\n                            ,TB1.FIFNIDCIAU ID_AGENCIA\r\n                            ,(select TRIM(FSGERAZSOC)  FROM PRODGRAL.GECCIAUN WHERE FIGEIDCIAU = TB1.FIFNIDCIAU AND FIGESTATUS = 1 ) RAZON_SOCIAL\r\n                            ,(select TRIM(FSGESIGCIA)  FROM PRODGRAL.GECCIAUN WHERE FIGEIDCIAU = TB1.FIFNIDCIAU AND FIGESTATUS = 1 ) FSGESIGCIA\r\n                            ,TRIM(FSFNRPTNM) REPORTE\r\n                            ,FIFNIDAGRP   IDGRUPO\r\n                            ,TRIM(FSFNIDAGRP) NOMBREGRUPOGENERAL \r\n                            ,FIFNIDGRPM  IDGRUPOMASTER\r\n                            ,TRIM(FSFNAGRPM)  NOMBREGRUPOMAESTRO\r\n                            ,INDEXGRUPO \r\n                            ,GRUPO  IDGRIPOGEN   \r\n                            ,TRIM(NOMBREGRUPO) NOMBREGRUPOAGRUPADOR\r\n                            ,TB1.ID_CONCEPTO\r\n                            ,TRIM(TB1.CONCEPTO) NOMBRE_CONCEPTO\r\n                            ,TB1.IDCUENTA\r\n                            ,SUBSTR(TB1.DESC_CTA,1,LOCATE(' ', TB1.DESC_CTA)) Key_CTA\r\n                            ,trim(TB1.DESC_CTA) DESCRIPCION_CUENTA\r\n                            ,(ifnull(CTAS.FDCOSALDOF,0)) Saldo_CUENTA\r\n                            ,CARTERAS.FICCIDCART\r\n                            ,TRIM(CARDESC.FSCCCVECAR) CLAVE_CARTERA\r\n                            ,TRIM(CARDESC.FSCCDESCAR) DESCRIPCION_CARTERA\r\n                            ,CASE WHEN CARTERAS.FICCIDCART = 0 THEN 0 ELSE COALESCE(CARTERAS.FDCCSALDOF,0) END Saldo_CARTERA\r\n                            FROM (\r\n                            SELECT DISTINCT \r\n\r\n                            0 FIFNIDDEP,\r\n                            PERI.FDCOIDPERI,\r\n                            PERI.FICOMES,\r\n                            PERI.FICOANIO,\r\n                            CONFIGREP.FIFNIDCIAU,\r\n                            CONFIGREP.FIFNIDRPT,\r\n                            REPORTES.FSFNRPTNM,\r\n                            CONFIGREP.FIFNIDAGRP,\r\n                            AGRPGENERAL.FSFNIDAGRP,\r\n                            CONFIGREP.FIFNIDGRPM,\r\n                            AGRPMAESTRO.FSFNAGRPM,\r\n                            AGRUPADOR.FIFNINDEX AS INDEXGRUPO,\r\n                            CONFIGREP.FIFNIDGRP AS GRUPO,\r\n                            AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,\r\n                            CONFIGREP.FIFNIDCPT AS ID_CONCEPTO,\r\n                            CONCEPTOS.FSFNNCPT AS CONCEPTO,\r\n                            COALESCE(CONFIG.FIFNIDCART,0) IDCART,\r\n                            COALESCE(CONFIG.FIFNIDCTA,0) IDCUENTA,\r\n                            (SELECT trim(CUENTA.FSCOCUENTA) || ' ' || CUENTA.FSCODESCTA FROM PRODCONT.COCATCTS CUENTA WHERE CUENTA.FICOIDTCCT = 1 AND cuenta.FICOIDCTA = CONFIG.FIFNIDCTA FETCH FIRST 1 ROW ONLY) DESC_CTA\r\n\r\n                            FROM \r\n                            PRODCONT.COCPERIO PERI, --PRODCONT.COCPERIO periodos\r\n                            PRODFINA.FNDCFREPCP CONFIGREP --PRODFINA.FNDCFREPCP Configuración del reporte cpt sus agrupadores\r\n                            LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP --PRODFINA.FNCAGRGRAL Agrupador general\r\n                            LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM --PRODFINA.FNCAGRUPA Agrupador maestro\r\n                            LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP --PRODFINA.FNCGRPCT Grupo o AGRUPADOR\r\n                            LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT AND CONCEPTOS.FIFNSTATUS = 1 --PRODFINA.FNCCONCP Conceptos\r\n                            LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT --PRODFINA.FNCREPORT Catalogo de reportes\r\n                            LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNIDTCCT IN (1) AND CONFIG.FIFNSTATUS = 1\r\n                            WHERE \r\n                            PERI.FICOANIO = " + anio + "\r\n                            --AND PERI.FICOMES = 12\r\n                            and REPORTES.FIFNIDRPT IN (4)\r\n--                            AND CONFIGREP.FIFNIDCIAU in (28) \r\n                            AND CONCEPTOS.FIFNSTATUS = 1\r\n                            AND REPORTES.FIFNSTATUS = 1\r\n                            AND CONFIG.FIFNSTATUS = 1\r\n                            ORDER BY \r\n                            CONFIGREP.FIFNIDRPT,\r\n                            CONFIGREP.FIFNIDCPT \r\n                            ) TB1 \r\n                            LEFT JOIN PRODCXC.CCESLCAS CARTERAS ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU \r\n                            AND CARTERAS.FICCIDCART = TB1.IDCART \r\n                            AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI\r\n                            AND CARTERAS.FICCSTATUS = 1\r\n                            LEFT join PRODCXC.CCCARTER CARDESC on CARTERAS.FICCIDCART = CARDESC.FICCIDCART AND CARDESC.FICCSTATUS = 1\r\n\r\n\r\n                            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI \r\n                            AND CTAS.FICOIDTCCT IN (1) \r\n                            AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU \r\n                            AND CTAS.FICOIDCTA = TB1.IDCUENTA \r\n                            AND CTAS.FICOSTATUS = 1 \r\n\r\n\r\n                            ) T1\r\n\r\n\r\n                            group BY\r\n                            ID_AGENCIA\r\n\r\n                            ,IDGRUPO\r\n                            ,NOMBREGRUPOGENERAL\r\n\r\n                            ,IDGRUPOMASTER\r\n                            ,NOMBREGRUPOMAESTRO\r\n\r\n                            ,INDEXGRUPO\r\n                            ,IDGRIPOGEN\r\n                            ,NOMBREGRUPOAGRUPADOR\r\n                            ,ID_CONCEPTO\r\n                            ,NOMBRE_CONCEPTO\r\n\r\n                            ORDER BY \r\n                            T1.IDGRUPO ASC";

            return dbCnx.GetDataTable(query);
        }

        public void GeneraExcelVP(List<int> idsAgencia)
        {
            List<Agencia> agencias;

            if (idsAgencia == null)
            {
                List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1);
                List<int> aIdAgencias = new List<int>();
                aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                agencias = Agencia.ListarPorIds(_db, aIdAgencias);
            }
            else
                agencias = Agencia.ListarPorIds(_db, idsAgencia);

            List<ConceptosContables> conceptosV1 = ConceptosContables.ListarVP(_db);



            DataTable dtVPExcelV1 = GetVPExcelV1();
            DataTable dtVPExcelV2 = GetVPExcelV2();
            DataTable dtVPWebV1 = new DataTable();
            DataTable dtVPWebV2 = new DataTable();


            int r = 1;
            int c = 1;
            int vExcel = 0;
            int vWeb = 0;
            int intAux = 0;
            DataRow[] drVPExcelV1 = null;
            DataRow[] drVPExcelV2 = null;
            DataRow[] drVPWebV1 = null;
            DataRow[] drVPWebV2 = null;

            List<ReporteVP> listaR = new List<ReporteVP>();
            List<ReporteGlobal> listaRG = new List<ReporteGlobal>();
            ReporteGlobal rG;

            using (DVAExcel.ExcelWriter eW = new DVAExcel.ExcelWriter(ruta))
            {
                foreach (Agencia agencia in agencias.OrderBy(o => o.Siglas))
                {
                    if ((agencia.Id == 590) || (agencia.Id == 583) || (agencia.Id == 301) || (agencia.Id == 583) || (agencia.Id == 563) || (agencia.Id == 593)
                       || (agencia.Id == 592) || (agencia.Id == 594) || (agencia.Id == 100))
                        continue;

                    //if (agencia.Id != 27 && agencia.Id != 36 && agencia.Id!=12)
                    //{



                    Console.WriteLine("[SIGLAS]: " + agencia.Siglas);

                    dtVPWebV1 = GetVPWebV1(agencia.Id, anio, mes);
                    dtVPWebV2 = GetVPWebV2(agencia.Id, anio, mes);

                    foreach (ConceptosContables concepto in conceptosV1)
                    {
                        ReporteVP rep = new ReporteVP();

                        r++;

                        rep.ID_CONCEPTO = concepto.Id;
                        rep.CONCEPTO = concepto.NombreConcepto;

                        drVPExcelV1 = dtVPExcelV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);

                        drVPWebV1 = dtVPWebV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);

                        drVPExcelV2 = dtVPExcelV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drVPWebV2 = dtVPWebV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);

                        if (drVPExcelV1.Length != 0)
                        {

                            vExcel = Convert.ToInt32(drVPExcelV1[0]["VALOR"]);



                            rep.EXCEL_V1 = vExcel;

                        }

                        if (drVPWebV1.Length != 0)
                        {
                            vWeb = 0;
                            decimal aa = 0;
                            decimal aaa = 0;
                            foreach (DataRow dr in drVPWebV1)
                            {

                                if ((concepto.Id == 107) || (concepto.Id == 109) || (concepto.Id == 110) || (concepto.Id == 111) || (concepto.Id == 112) || (concepto.Id == 113) || (concepto.Id == 114) || (concepto.Id == 115) || (concepto.Id == 116))
                                {

                                    vWeb += Convert.ToInt32(Convert.ToDecimal(dr["VALOR"]) * 100);
                                }
                                else
                                {
                                    vWeb += Convert.ToInt32(dr["VALOR"]);
                                }
                            }
                            rep.WEB_V1 = vWeb;

                            vWeb = 0;

                        }

                        rep.DIFF_V1 = rep.EXCEL_V1 - rep.WEB_V1;


                        if (drVPExcelV2.Length != 0)
                        {
                            vExcel = Convert.ToInt32(drVPExcelV2[0]["VALOR"]);
                            rep.EXCEL_V2 = vExcel;
                        }

                        if (drVPWebV2.Length != 0)
                        {
                            vWeb = 0;

                            foreach (DataRow dr in drVPWebV2)
                            {

                                if ((concepto.Id == 107) || (concepto.Id == 109) || (concepto.Id == 110) || (concepto.Id == 111) || (concepto.Id == 112) || (concepto.Id == 113) || (concepto.Id == 114) || (concepto.Id == 115) || (concepto.Id == 116))
                                {

                                    vWeb += Convert.ToInt32(Convert.ToDecimal(dr["VALOR"]) * 100);
                                }
                                else
                                {
                                    vWeb += Convert.ToInt32(dr["VALOR"]);

                                }
                            }
                            rep.WEB_V2 = vWeb;

                            vWeb = 0;

                        }

                        rep.DIFF_V2 = rep.EXCEL_V2 - rep.WEB_V2;


                        Console.WriteLine("[CONCEPTO]: " + concepto.Id + "_" + concepto.NombreConcepto +
                            " [EXCEL_V1]: " + rep.EXCEL_V1 +
                            " [WEB_V1]: " + rep.WEB_V1 +
                            " [DIFF_V1]: " + rep.DIFF_V1 +
                            " [EXCEL_V2]: " + rep.EXCEL_V2 +
                            " [WEB_V2]: " + rep.WEB_V2 +
                            " [DIFF_V2]: " + rep.DIFF_V2);


                        listaR.Add(rep);
                    }

                    eW.WriteDataTable(listaR.ToDataTable(), agencia.Siglas);
                    
                    listaR = new List<ReporteVP>();
                    //}
                }

                eW.Dispose();
            }            
        }


        public DataTable GetVPExcelV1()
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RESMEX . FIFNYEAR ANIO, RESMEX . FIFNMONTH MES, RESMEX . FIFNVALUE VALOR, CONCP . FIFNAGRPM ORDEN\r\n" +
                "FROM PRODFINA.FNDRESMEX RESMEX\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RESMEX . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RESMEX . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RESMEX . FIFNYEAR = " + anio + " AND RESMEX . FIFNMONTH = " + mes + " AND RESMEX . FIFNIDRPT = 9";

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetVPExcelV2()
        {
            query = "SELECT \r\n" +
                "CIAUN . FIGEIDCIAU ID_AGENCIA, TRIM(CIAUN . FSGERAZSOC) RAZON_SOCIAL, CONCAT(CONCAT(CIAUN . FIGEIDCIAU, '_'), TRIM(CIAUN . FSGERAZSOC)) EMPRESA, " +
                "TRIM(CIAUN . FSGESIGCIA) SIGLAS,\r\n" +
                "CONCP . FIFNIDCPT ID_CONCEPTO, TRIM(CONCP . FSFNNCPT) DESC_CONCEPTO, CONCAT(CONCAT(CONCP . FIFNIDCPT, '_'), TRIM(CONCP . FSFNNCPT)) CONCEPTO, \r\n" +
                "RSMEEX . FIFNYEAR ANIO, RSMEEX . FIFNMONTH MES, RSMEEX . FIFNVALUE VALOR, CONCP . FIFNORDEV2 ORDEN\r\n" +
                "FROM PRODFINA.FNDRSMEEX RSMEEX\r\n" +
                "INNER JOIN PRODGRAL . GECCIAUN CIAUN ON RSMEEX . FIFNIDCIAU = CIAUN . FIGEIDCIAU\r\n" +
                "INNER JOIN PRODFINA . FNCCONCP CONCP ON RSMEEX . FIFNCPTD = CONCP . FIFNIDCPT\r\n" +
                "WHERE RSMEEX . FIFNYEAR = " + anio + " AND RSMEEX . FIFNMONTH = " + mes + " AND RSMEEX . FIFNIDRPT = 9";

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetVPWebV1(int aIdAgencia, int aAnio, int aMes)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("ID_AGENCIA", typeof(int));
            dataTable.Columns.Add("ID_CONCEPTO", typeof(int));
            dataTable.Columns.Add("VALOR", typeof(decimal));

            List<VolumenYPorcentaje> ListaResultados = VolumenYPorcentaje.Listar(_db, aIdAgencia, aAnio);

            foreach (var Registro in ListaResultados)
            {
                DataRow row = dataTable.NewRow();
                row["ID_AGENCIA"] = aIdAgencia;
                row["ID_CONCEPTO"] = Registro.IdConcept;

                switch (aMes)
                {
                    case 1:
                        row["VALOR"] = Registro.Ene;
                        break;
                    case 2:
                        row["VALOR"] = Registro.Feb;
                        break;
                    case 3:
                        row["VALOR"] = Registro.Mar;
                        break;
                    case 4:
                        row["VALOR"] = Registro.Abr;
                        break;
                    case 5:
                        row["VALOR"] = Registro.May;
                        break;
                    case 6:
                        row["VALOR"] = Registro.Jun;
                        break;
                    case 7:
                        row["VALOR"] = Registro.Jul;
                        break;
                    case 8:
                        row["VALOR"] = Registro.Ago;
                        break;
                    case 9:
                        row["VALOR"] = Registro.Sep;
                        break;
                    case 10:
                        row["VALOR"] = Registro.Oct;

                        break;
                    case 11:
                        row["VALOR"] = Registro.Nov;

                        break;
                    case 12:
                        row["VALOR"] = Registro.Dic;

                        break;

                }



                dataTable.Rows.Add(row);

            }



            return dataTable;
        }

        public DataTable GetVPWebV2(int aIdAgencia, int aAnio, int aMes)
        {
            // Crear un DataTable
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("ID_AGENCIA", typeof(int));
            dataTable.Columns.Add("ID_CONCEPTO", typeof(int));
            dataTable.Columns.Add("VALOR", typeof(decimal));

            List<VolumenYPorcentaje> ListaResultados = VolumenYPorcentaje.ListarExtralibros(_db, aIdAgencia, aAnio);

            foreach (var Registro in ListaResultados)
            {
                DataRow row = dataTable.NewRow();
                row["ID_AGENCIA"] = aIdAgencia;
                row["ID_CONCEPTO"] = Registro.IdConcept;

                switch (aMes)
                {
                    case 1:
                        row["VALOR"] = Registro.Ene;
                        break;
                    case 2:
                        row["VALOR"] = Registro.Feb;
                        break;
                    case 3:
                        row["VALOR"] = Registro.Mar;
                        break;
                    case 4:
                        row["VALOR"] = Registro.Abr;
                        break;
                    case 5:
                        row["VALOR"] = Registro.May;
                        break;
                    case 6:
                        row["VALOR"] = Registro.Jun;
                        break;
                    case 7:
                        row["VALOR"] = Registro.Jul;
                        break;
                    case 8:
                        row["VALOR"] = Registro.Ago;
                        break;
                    case 9:
                        row["VALOR"] = Registro.Sep;
                        break;
                    case 10:
                        row["VALOR"] = Registro.Oct;

                        break;
                    case 11:
                        row["VALOR"] = Registro.Nov;

                        break;
                    case 12:
                        row["VALOR"] = Registro.Dic;

                        break;

                }



                dataTable.Rows.Add(row);

            }



            return dataTable;
        }

        public void GeneraExcelCuentas(List<int> idsAgencia, bool esUnaAgencia) // EDITAR  ---------------------------------------------------------------------------------------------------------
        {
            List<Agencia> agencias;

            if (idsAgencia == null)
            {
                List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1); //Todas las agencias
                List<int> aIdAgencias = new List<int>();

                aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                agencias = Agencia.ListarPorIds(_db, aIdAgencias);
            }
            else
                agencias = Agencia.ListarPorIds(_db, idsAgencia);

            string idAgenciaString = string.Join(", ", agencias.Select(item => item.Id));

            List<ConceptosContables> conceptosV1 = ConceptosContables.ListarBGV1yV2(_db); //obtiene el listado de todos los conceptos contables
            //List<ConceptosContables> conceptosV2 = ConceptosContables.ListarBGV1yV2(_db);
            List<ConceptosContables> conceptosV2 = ConceptosContables.ListarCuentasV1yV2(_db);

            DataTable dtBGExcelV1 = GetExcelV1(8);
            DataTable dtBGExcelV1Acum = GetExcelV1Acumulado(8);
            DataTable dtBGExcelV2 = GetExcelV2(8);
            DataTable dtBGExcelV2Acum = GetExcelV2Acumulado(8);

            //DataTable dtBGWebV2Acum = GetBGWebV1();

            int r = 1;
            int c = 1;
            int vExcel = 0;
            int vWeb = 0;
            int intAux = 0;
            DataRow[] drBGExcelV1 = null;
            DataRow[] drBGExcelV1Acum = null;
            DataRow[] drBGExcelV2 = null;
            DataRow[] drBGExcelV2Acum = null;
            DataRow[] drBGWebV1 = null;
            DataRow[] drBGWebV1Acum = null;
            DataRow[] drBGWebV2 = null;
            DataRow[] drBGWebV2Acum = null;

            List<Reporte> listaR = new List<Reporte>();
            List<ReporteGlobal> listaRG = new List<ReporteGlobal>();
            ReporteGlobal rG;


            DataTable dtBGWebV1 = new DataTable();
            DataTable dtBGWebV2 = new DataTable();


            using (DVAExcel.ExcelWriter eW = new DVAExcel.ExcelWriter(ruta))
            {
                foreach (Agencia agencia in agencias.OrderBy(o => o.Siglas))
                {
                    if (esUnaAgencia)
                    {
                        //Metodos a iterar ----------------------------------        
                        dtBGWebV1 = GetCuentasWebV1(idAgenciaString);

                        dtBGWebV2 = GetCuentasWebV2(idAgenciaString);
                    }
                    else
                    {
                        dtBGWebV1 = GetCuentasWebV1(agencia.Id.ToString());

                        dtBGWebV2 = GetCuentasWebV2(agencia.Id.ToString());
                    }



                    if ((agencia.Id == 590) || (agencia.Id == 583) || (agencia.Id == 301) || (agencia.Id == 583) || (agencia.Id == 563) || (agencia.Id == 593)
                        || (agencia.Id == 592) || (agencia.Id == 594) || (agencia.Id == 100))
                        continue;

                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("[SIGLAS]: " + agencia.Siglas);
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");

                    foreach (ConceptosContables concepto in conceptosV2)
                    {
                        Reporte rep = new Reporte();

                        r++;

                        rep.ID_CONCEPTO = concepto.Id;
                        rep.CONCEPTO = concepto.NombreConcepto;

                        int cambioSigo = 1;
                        //if ((concepto.Id == 709) || (concepto.Id == 710) || (concepto.Id == 711) || (concepto.Id == 712) || (concepto.Id == 713) ||
                        //    (concepto.Id == 714) || (concepto.Id == 715) || (concepto.Id == 716) || (concepto.Id == 717) || (concepto.Id == 718) ||
                        //    (concepto.Id == 719) || (concepto.Id == 720) || (concepto.Id == 721) || (concepto.Id == 722) || (concepto.Id == 723) ||
                        //    (concepto.Id == 724) || (concepto.Id == 725) || (concepto.Id == 726) || (concepto.Id == 727) || (concepto.Id == 728) ||
                        //    (concepto.Id == 729) || (concepto.Id == 730) || (concepto.Id == 731) || (concepto.Id == 732) || (concepto.Id == 733) ||
                        //    (concepto.Id == 734) || (concepto.Id == 735) || (concepto.Id == 736) || (concepto.Id == 737) || (concepto.Id == 738) ||
                        //    (concepto.Id == 739) || (concepto.Id == 740) || (concepto.Id == 741) || (concepto.Id == 742) || (concepto.Id == 743) ||
                        //    (concepto.Id == 744) || (concepto.Id == 745) || (concepto.Id == 746) || (concepto.Id == 747) || (concepto.Id == 748) ||
                        //    (concepto.Id == 749) || (concepto.Id == 750) || (concepto.Id == 751) || (concepto.Id == 752) || (concepto.Id == 753) ||
                        //    (concepto.Id == 754) || (concepto.Id == 755) || (concepto.Id == 756) || (concepto.Id == 757) || (concepto.Id == 758) ||
                        //    (concepto.Id == 759) || (concepto.Id == 760) || (concepto.Id == 761))
                        //    cambioSigo = -1;

                        // drBGExcelV1 = dtBGExcelV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drBGExcelV1 = dtBGExcelV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drBGExcelV1Acum = dtBGExcelV1Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        if (agencia.Id == 27)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 286)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 36)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 12)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 35)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 588)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 32)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 212)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }

                        drBGExcelV2 = dtBGExcelV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drBGExcelV2Acum = dtBGExcelV2Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        if (agencia.Id == 27)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 286)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 36)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 12)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 35)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 588)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 32)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 212)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }

                        if (drBGExcelV1.Length != 0)
                        {
                            vExcel = Convert.ToInt32(drBGExcelV1[0]["VALOR"]);
                            rep.EXCEL_V1 = vExcel;
                            vExcel = Convert.ToInt32(drBGExcelV1Acum[0]["VALOR"]);
                            rep.EXCEL_V1_ACUM = vExcel;
                        }

                        if (drBGWebV1.Length != 0)
                        {
                            vWeb = 0;

                            foreach (DataRow dr in drBGWebV1)
                            {
                                if (mes == 1)
                                    vWeb += Convert.ToInt32(dr["ENERO"]);
                                else if (mes == 2)
                                    vWeb += Convert.ToInt32(dr["FEBRERO"]);
                                else if (mes == 3)
                                    vWeb += Convert.ToInt32(dr["MARZO"]);
                                else if (mes == 4)
                                    vWeb += Convert.ToInt32(dr["ABRIL"]);
                                else if (mes == 5)
                                    vWeb += Convert.ToInt32(dr["MAYO"]);
                                else if (mes == 6)
                                    vWeb += Convert.ToInt32(dr["JUNIO"]);
                                else if (mes == 7)
                                    vWeb += Convert.ToInt32(dr["JULIO"]);
                                else if (mes == 8)
                                    vWeb += Convert.ToInt32(dr["AGOSTO"]);
                                else if (mes == 9)
                                    vWeb += Convert.ToInt32(dr["SEPTIEMBRE"]);
                                else if (mes == 10)
                                    vWeb += Convert.ToInt32(dr["OCTUBRE"]);
                                else if (mes == 11)
                                    vWeb += Convert.ToInt32(dr["NOVIEMBRE"]);
                                else if (mes == 12)
                                    vWeb += Convert.ToInt32(dr["DICIEMBRE"]);

                                rep.WEB_V1 = vWeb * cambioSigo / 1000;
                            }

                            vWeb = 0;

                            foreach (DataRow dr in drBGWebV1Acum)
                            {
                                if (mes == 1)
                                    vWeb += Convert.ToInt32(dr["ENERO"]);
                                else if (mes == 2)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]);
                                else if (mes == 3)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]);
                                else if (mes == 4)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]);
                                else if (mes == 5)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]);
                                else if (mes == 6)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]);
                                else if (mes == 7)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]);
                                else if (mes == 8)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]);
                                else if (mes == 9)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]) +
                                        Convert.ToInt32(dr["SEPTIEMBRE"]);
                                else if (mes == 10)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]) +
                                        Convert.ToInt32(dr["SEPTIEMBRE"]) + Convert.ToInt32(dr["OCTUBRE"]);
                                else if (mes == 11)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]) +
                                        Convert.ToInt32(dr["SEPTIEMBRE"]) + Convert.ToInt32(dr["OCTUBRE"]) +
                                        Convert.ToInt32(dr["NOVIEMBRE"]);
                                else if (mes == 12)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]) +
                                        Convert.ToInt32(dr["SEPTIEMBRE"]) + Convert.ToInt32(dr["OCTUBRE"]) +
                                        Convert.ToInt32(dr["NOVIEMBRE"]) + Convert.ToInt32(dr["DICIEMBRE"]);

                                rep.WEB_V1_ACUM = vWeb * cambioSigo / 1000;
                            }
                        }

                        rep.DIFF_V1 = rep.EXCEL_V1 - rep.WEB_V1;
                        rep.DIFF_V1_ACUM = rep.EXCEL_V1_ACUM - rep.WEB_V1_ACUM;

                        if (drBGExcelV2.Length != 0)
                        {
                            vExcel = Convert.ToInt32(drBGExcelV2[0]["VALOR"]);
                            rep.EXCEL_V2 = vExcel;
                            vExcel = Convert.ToInt32(drBGExcelV2Acum[0]["VALOR"]);
                            rep.EXCEL_V2_ACUM = vExcel;
                        }

                        if (drBGWebV2.Length != 0)
                        {
                            vWeb = 0;

                            foreach (DataRow dr in drBGWebV2)
                            {
                                if (mes == 1)
                                    vWeb += Convert.ToInt32(dr["ENERO"]);
                                else if (mes == 2)
                                    vWeb += Convert.ToInt32(dr["FEBRERO"]);
                                else if (mes == 3)
                                    vWeb += Convert.ToInt32(dr["MARZO"]);
                                else if (mes == 4)
                                    vWeb += Convert.ToInt32(dr["ABRIL"]);
                                else if (mes == 5)
                                    vWeb += Convert.ToInt32(dr["MAYO"]);
                                else if (mes == 6)
                                    vWeb += Convert.ToInt32(dr["JUNIO"]);
                                else if (mes == 7)
                                    vWeb += Convert.ToInt32(dr["JULIO"]);
                                else if (mes == 8)
                                    vWeb += Convert.ToInt32(dr["AGOSTO"]);
                                else if (mes == 9)
                                    vWeb += Convert.ToInt32(dr["SEPTIEMBRE"]);
                                else if (mes == 10)
                                    vWeb += Convert.ToInt32(dr["OCTUBRE"]);
                                else if (mes == 11)
                                    vWeb += Convert.ToInt32(dr["NOVIEMBRE"]);
                                else if (mes == 12)
                                    vWeb += Convert.ToInt32(dr["DICIEMBRE"]);

                                rep.WEB_V2 = vWeb * cambioSigo;
                            }

                            vWeb = 0;

                            foreach (DataRow dr in drBGWebV2Acum)
                            {
                                if (mes == 1)
                                    vWeb += Convert.ToInt32(dr["ENERO"]);
                                else if (mes == 2)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]);
                                else if (mes == 3)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]);
                                else if (mes == 4)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]);
                                else if (mes == 5)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]);
                                else if (mes == 6)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]);
                                else if (mes == 7)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]);
                                else if (mes == 8)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]);
                                else if (mes == 9)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]) +
                                        Convert.ToInt32(dr["SEPTIEMBRE"]);
                                else if (mes == 10)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]) +
                                        Convert.ToInt32(dr["SEPTIEMBRE"]) + Convert.ToInt32(dr["OCTUBRE"]);
                                else if (mes == 11)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]) +
                                        Convert.ToInt32(dr["SEPTIEMBRE"]) + Convert.ToInt32(dr["OCTUBRE"]) +
                                        Convert.ToInt32(dr["NOVIEMBRE"]);
                                else if (mes == 12)
                                    vWeb += Convert.ToInt32(dr["ENERO"]) + Convert.ToInt32(dr["FEBRERO"]) +
                                        Convert.ToInt32(dr["MARZO"]) + Convert.ToInt32(dr["ABRIL"]) +
                                        Convert.ToInt32(dr["MAYO"]) + Convert.ToInt32(dr["JUNIO"]) +
                                        Convert.ToInt32(dr["JULIO"]) + Convert.ToInt32(dr["AGOSTO"]) +
                                        Convert.ToInt32(dr["SEPTIEMBRE"]) + Convert.ToInt32(dr["OCTUBRE"]) +
                                        Convert.ToInt32(dr["NOVIEMBRE"]) + Convert.ToInt32(dr["DICIEMBRE"]);

                                rep.WEB_V2_ACUM = vWeb * cambioSigo;
                            }
                        }

                        rep.DIFF_V2 = rep.EXCEL_V2 - rep.WEB_V2;
                        rep.DIFF_V2_ACUM = rep.EXCEL_V2_ACUM - rep.WEB_V2_ACUM;

                        Console.WriteLine("[CONCEPTO]: " + concepto.Id + "_" + concepto.NombreConcepto +
                            " [EXCEL_V1]: " + rep.EXCEL_V1 +
                            " [WEB_V1]: " + rep.WEB_V1 +
                            " [DIFF_V1]: " + rep.DIFF_V1 +
                            " [EXCEL_V1_ACUMULADO]: " + rep.EXCEL_V1_ACUM +
                            " [WEB_V1_ACUMULADO]: " + rep.WEB_V1_ACUM +
                            " [DIFF_V1_ACUMULADO]: " + rep.DIFF_V1_ACUM +
                            " [EXCEL_V2]: " + rep.EXCEL_V2 +
                            " [WEB_V2]: " + rep.WEB_V2 +
                            " [DIFF_V2]: " + rep.DIFF_V2 +
                            " [EXCEL_V2_ACUMULADO]: " + rep.EXCEL_V2_ACUM +
                            " [WEB_V2_ACUMULADO]: " + rep.WEB_V2_ACUM +
                            " [DIFF_V2_ACUMULADO]: " + rep.DIFF_V2_ACUM);

                        listaR.Add(rep);
                    }

                    var dtToExcel = listaR.ToDataTable();
                    eW.WriteDataTable(listaR.ToDataTable(), agencia.Siglas);
                    eW.Dispose();

                    listaR = new List<Reporte>();
                }
            }
        }

        public DataTable GetCuentasWebV1(string Id_Agencia)
        {
            query = @"SELECT 
        GRUPO, INDEX,
            NOMBREGRUPO,
            IDCONCEPTO,
            CONCEPTO,
            SUM(Enero) Enero,
            SUM(Febrero) Febrero,
            SUM(Marzo) Marzo,
            SUM(Abril) Abril,
            SUM(Mayo) Mayo,
            SUM(Junio) Junio,
            SUM(Julio) Julio,
            SUM(Agosto) Agosto,
            SUM(Septiembre) Septiembre,
            SUM(Octubre) Octubre,
            SUM(Noviembre) Noviembre,
            SUM(Diciembre) Diciembre,
            SUM(Total) Total,
            CUENTAS,
            CARTERAS
            FROM (
            SELECT DISTINCT
            TB1.GRUPO,TB1.INDEX, 
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  1  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Enero,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  2  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Febrero,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  3  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Marzo,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  4  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Abril,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  5  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Mayo,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  6  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Junio,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  7  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Julio,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  8  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Agosto,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  9  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Septiembre,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  10 THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Octubre,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  11 THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Noviembre,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  12 THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Diciembre,
            ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END))),0),0) AS Total,
	        (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1 AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS
	
        FROM (
            SELECT DISTINCT 
        
                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM,CONFIGREP.FIFNIDGRP AS GRUPO,  AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART,0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA,0) IDCUENTA
            FROM 
                PRODCONT.COCPERIO PERI, 
                PRODFINA.FNDCFREPCP CONFIGREP 
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP 
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM  
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP  
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1 
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT 
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE 
                PERI.FICOANIO = " + anio + @" AND  
               CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN (" + Id_Agencia + ')' + @"

                AND CONCEPTOS.FIFNIDCPT <> 716



            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN (1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO


        UNION

        SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0), 0) AS Enero,
                ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Febrero,
                   ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Marzo,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Abril,
                         ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Mayo,
                            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Junio,
                               ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Julio,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Agosto,
                                     ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Septiembre,
                                        ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Octubre,
                                           ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Noviembre,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Diciembre,
                                                 ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END))), 0),0) AS Total,
                                          
                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                          
                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                          
                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' +

               @" AND CONCEPTOS.FIFNIDCPT IN(716)



            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 71100 IDCONCEPTO, 'Contratos en Transito' CONCEPTO,
        SUM(Enero) Enero,
        SUM(Febrero) Febrero,
        SUM(Marzo) Marzo,
        SUM(Abril) Abril,
        SUM(Mayo) Mayo,
        SUM(Junio) Junio,
        SUM(Julio) Julio,
        SUM(Agosto) Agosto,
        SUM(Septiembre) Septiembre,
        SUM(Octubre) Octubre,
        SUM(Noviembre) Noviembre,
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Enero,
            IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Febrero,
                IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Marzo,
                    IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Abril,
                        IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Mayo,
                            IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Junio,
                                IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Julio,
                                    IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Agosto,
                                        IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Septiembre,
                                            IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Octubre,
                                                IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Noviembre,
                                                    IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Diciembre,
                                                        IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END))), 0) AS Total,
                                                 
                                                             (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                 
                                                             LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                 
                                                             WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 31

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 71800 IDCONCEPTO, 'Utilidad Bruta Mano de Obra Taller' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 32 AND IDCONCEPTO IN(713, 714, 715, 716, 717)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 72100 IDCONCEPTO, 'Utilidad Bruta Mano de Obra Garantias' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 33 AND IDCONCEPTO IN(718, 719, 720)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 72400 IDCONCEPTO, 'Utilidad Bruta T.O.T.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                

            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 34 AND IDCONCEPTO IN(721, 722, 723)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 72700 IDCONCEPTO, 'Utilidad Bruta Materiales Diversos.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 35 AND IDCONCEPTO IN(724, 725, 726)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 73400 IDCONCEPTO, 'Utilidad Bruta Refacciones Mostrador.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 36 AND IDCONCEPTO IN(731, 732, 733)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 73700 IDCONCEPTO, 'Utilidad Bruta Refacciones Taller.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 37 AND IDCONCEPTO IN(734, 735, 736)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 74000 IDCONCEPTO, 'Utilidad Bruta Refacciones Cía. de Seguro.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 38 AND IDCONCEPTO IN(737, 738, 739)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 74300 IDCONCEPTO, 'Utilidad Bruta Refacciones Garantía.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 39 AND IDCONCEPTO IN(740, 741, 742)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 74600 IDCONCEPTO, 'Utilidad Bruta Refacciones Otras Mercancias.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 40 AND IDCONCEPTO IN(743, 744, 745)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 74900 IDCONCEPTO, 'Utilidad Bruta Refacciones Accesorios.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 41 AND IDCONCEPTO IN(746, 747, 748)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 75400 IDCONCEPTO, 'Utilidad Bruta Hojalateria y Pintura.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 42 AND IDCONCEPTO IN(749, 750, 751, 752, 753)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 76300 IDCONCEPTO, 'Utilidad Bruta Materiales Diversos Hojalateria y Pintura.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 43 AND IDCONCEPTO IN(754, 755, 762)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION

        SELECT 45 GRUPO,15 INDEX, 'o - Utilidad Bruta Servicio' NOMBREGRUPO, 75700 IDCONCEPTO, 'Utilidad Bruta Servicio' CONCEPTO,  
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 1 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Enero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 2 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Febrero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 3 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Marzo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 4 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Abril,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 5 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Mayo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 6 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Junio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 7 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Julio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 8 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Agosto,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 9 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Septiembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 10 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Octubre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 11 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Noviembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Diciembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES <= 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Total
        , '' CUENTAS, '' CARTERAS
        FROM(
            SELECT FIFNMONTH FICOMES, FIFNVALUE Utilidad_Bruta_Servicio FROM PRODFINA.FNDRESMEN
                WHERE FIFNYEAR = " + anio + @"
                AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                AND FIFNCPTD = 1022
       )
        UNION


        SELECT 46 GRUPO, 16 INDEX, 'p - Utilidad Bruta Refacciones' NOMBREGRUPO, 76100 IDCONCEPTO, 'Utilidad Bruta Refacciones' CONCEPTO,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 1 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Enero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 2 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Febrero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 3 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Marzo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 4 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Abril,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 5 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Mayo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 6 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Junio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 7 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Julio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 8 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Agosto,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 9 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Septiembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 10 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Octubre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 11 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Noviembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Diciembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES <= 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Total
        , '' CUENTAS
        , '' CARTERAS
        FROM(
               SELECT FIFNMONTH FICOMES, FIFNVALUE Utilidad_Bruta_Servicio FROM PRODFINA.FNDRESMEN
                WHERE FIFNYEAR = " + anio + @"
                AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                AND FIFNCPTD = 1039
       )


        UNION


        SELECT 47 GRUPO, 17 INDEX, 'q - Utilidad Bruta Hojalateria y Pintura' NOMBREGRUPO, 76200 IDCONCEPTO, 'Utilidad Bruta Hojalateria y Pintura' CONCEPTO,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 1 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Enero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 2 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Febrero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 3 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Marzo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 4 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Abril,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 5 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Mayo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 6 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Junio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 7 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Julio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 8 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Agosto,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 9 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Septiembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 10 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Octubre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 11 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Noviembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Diciembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES <= 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Total
        , '' CUENTAS
        , '' CARTERAS
        FROM(
               SELECT FIFNMONTH FICOMES, FIFNVALUE Utilidad_Bruta_Servicio FROM PRODFINA.FNDRESMEN
                WHERE FIFNYEAR = " + anio + @"
                AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                AND FIFNCPTD = 1026
       )

       order by GRUPO, IDCONCEPTO
       )
        GROUP BY
            GRUPO, INDEX,
            NOMBREGRUPO,
            IDCONCEPTO,
            CONCEPTO,
            CUENTAS,
            CARTERAS


        ORDER BY INDEX";

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetCuentasWebV2(string Id_Agencia)
        {
            query = @"SELECT 
        GRUPO, INDEX,
            NOMBREGRUPO,
            IDCONCEPTO,
            CONCEPTO,
            SUM(Enero) Enero,
            SUM(Febrero) Febrero,
            SUM(Marzo) Marzo,
            SUM(Abril) Abril,
            SUM(Mayo) Mayo,
            SUM(Junio) Junio,
            SUM(Julio) Julio,
            SUM(Agosto) Agosto,
            SUM(Septiembre) Septiembre,
            SUM(Octubre) Octubre,
            SUM(Noviembre) Noviembre,
            SUM(Diciembre) Diciembre,
            SUM(Total) Total,
            CUENTAS,
            CARTERAS
            FROM (
            SELECT DISTINCT
            TB1.GRUPO,TB1.INDEX, 
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  1  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Enero,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  2  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Febrero,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  3  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Marzo,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  4  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Abril,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  5  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Mayo,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  6  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Junio,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  7  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Julio,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  8  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Agosto,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  9  THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Septiembre,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  10 THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Octubre,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  11 THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Noviembre,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES =  12 THEN CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END END) )),0),0) AS Diciembre,
            ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN (709,710,711,712) THEN (COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0)) ELSE (COALESCE(CTAS.FDCOTOTABO,0) - COALESCE(CTAS.FDCOTOTCAR,0)) END))),0),0) AS Total,
	        (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1 AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS
	
        FROM (
            SELECT DISTINCT 
        
                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM,CONFIGREP.FIFNIDGRP AS GRUPO,  AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART,0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA,0) IDCUENTA
            FROM 
                PRODCONT.COCPERIO PERI, 
                PRODFINA.FNDCFREPCP CONFIGREP 
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP 
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM  
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP  
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1 
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT 
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE 
                PERI.FICOANIO = " + anio + @" AND  
               CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN (" + Id_Agencia + ')' + @"

                AND CONCEPTOS.FIFNIDCPT <> 716



            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN (1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO


        UNION

        SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,
            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0), 0) AS Enero,
                ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Febrero,
                   ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Marzo,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Abril,
                         ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Mayo,
                            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Junio,
                               ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Julio,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Agosto,
                                     ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Septiembre,
                                        ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Octubre,
                                           ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Noviembre,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0),0) AS Diciembre,
                                                 ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END))), 0),0) AS Total,
                                          
                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                          
                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                          
                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' +

               @" AND CONCEPTOS.FIFNIDCPT IN(716)



            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 711 IDCONCEPTO, 'Contratos en Transito' CONCEPTO,
        SUM(Enero) Enero,
        SUM(Febrero) Febrero,
        SUM(Marzo) Marzo,
        SUM(Abril) Abril,
        SUM(Mayo) Mayo,
        SUM(Junio) Junio,
        SUM(Julio) Julio,
        SUM(Agosto) Agosto,
        SUM(Septiembre) Septiembre,
        SUM(Octubre) Octubre,
        SUM(Noviembre) Noviembre,
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Enero,
            IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Febrero,
                IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Marzo,
                    IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Abril,
                        IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Mayo,
                            IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Junio,
                                IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Julio,
                                    IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Agosto,
                                        IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Septiembre,
                                            IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Octubre,
                                                IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Noviembre,
                                                    IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) AS Diciembre,
                                                        IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END))), 0) AS Total,
                                                 
                                                             (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                 
                                                             LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                 
                                                             WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 31

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 718 IDCONCEPTO, 'Utilidad Bruta Mano de Obra Taller' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 32 AND IDCONCEPTO IN(713, 714, 715, 716, 717)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 721 IDCONCEPTO, 'Utilidad Bruta Mano de Obra Garantias' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 33 AND IDCONCEPTO IN(718, 719, 720)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 724 IDCONCEPTO, 'Utilidad Bruta T.O.T.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                

            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 34 AND IDCONCEPTO IN(721, 722, 723)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 727 IDCONCEPTO, 'Utilidad Bruta Materiales Diversos.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 35 AND IDCONCEPTO IN(724, 725, 726)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 734 IDCONCEPTO, 'Utilidad Bruta Refacciones Mostrador.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 36 AND IDCONCEPTO IN(731, 732, 733)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 737 IDCONCEPTO, 'Utilidad Bruta Refacciones Taller.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 37 AND IDCONCEPTO IN(734, 735, 736)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 740 IDCONCEPTO, 'Utilidad Bruta Refacciones Cía. de Seguro.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 38 AND IDCONCEPTO IN(737, 738, 739)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 743 IDCONCEPTO, 'Utilidad Bruta Refacciones Garantía.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 39 AND IDCONCEPTO IN(740, 741, 742)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 746 IDCONCEPTO, 'Utilidad Bruta Refacciones Otras Mercancias.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 40 AND IDCONCEPTO IN(743, 744, 745)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 749 IDCONCEPTO, 'Utilidad Bruta Refacciones Accesorios.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 41 AND IDCONCEPTO IN(746, 747, 748)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 754 IDCONCEPTO, 'Utilidad Bruta Hojalateria y Pintura.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 42 AND IDCONCEPTO IN(749, 750, 751, 752, 753)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION


        SELECT GRUPO, INDEX, NOMBREGRUPO, 763 IDCONCEPTO, 'Utilidad Bruta Materiales Diversos Hojalateria y Pintura.' CONCEPTO, 
        SUM(Enero) Enero, 
        SUM(Febrero) Febrero, 
        SUM(Marzo) Marzo, 
        SUM(Abril) Abril, 
        SUM(Mayo) Mayo, 
        SUM(Junio) Junio, 
        SUM(Julio) Julio, 
        SUM(Agosto) Agosto, 
        SUM(Septiembre) Septiembre, 
        SUM(Octubre) Octubre, 
        SUM(Noviembre) Noviembre, 
        SUM(Diciembre) Diciembre,
        SUM(Total) Total, '' CUENTAS, '' CARTERAS
        FROM(
            SELECT DISTINCT
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO,

            ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 1  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0) / 1000, 0) AS Enero,
                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 2  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Febrero,
                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 3  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Marzo,
                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 4  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Abril,
                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 5  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Mayo,
                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 6  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Junio,
                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 7  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Julio,
                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 8  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Agosto,
                                              ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 9  THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Septiembre,
                                                  ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 10 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Octubre,
                                                      ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 11 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Noviembre,
                                                          ROUND(IFNULL((SUM((CASE WHEN TB1.FICOMES = 12 THEN CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END END))), 0)/ 1000,0) AS Diciembre,
                                                              ROUND(IFNULL((SUM((CASE WHEN TB1.IDCONCEPTO IN(709, 710, 711, 712) THEN(COALESCE(CARTERAS.FDCCSALDOF, 0) + COALESCE(CTAS.FDCOSALDOF, 0)) ELSE(COALESCE(CTAS.FDCOTOTABO, 0) - COALESCE(CTAS.FDCOTOTCAR, 0)) END) * (-1))), 0)/ 1000,0) AS Total,
                                                          
                                                                      (SELECT LISTAGG(TRIM(CTA.FSCOCUENTA), ', ') FROM PRODFINA.FNDCFGCPT CFG
                                                          
                                                                      LEFT JOIN PRODCONT.COCATCTS CTA ON CTA.FICOIDCTA = CFG.FIFNIDCTA
                                                          
                                                                      WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CTA.FICOIDTCCT = 1  AND CFG.FIFNSTATUS = 1) CUENTAS,
            (SELECT LISTAGG(TRIM(CART.FSCCCVECAR), ', ') FROM PRODFINA.FNDCFGCPT CFG
            LEFT JOIN PRODCXC.CCCARTER CART ON CART.FICCIDCART = CFG.FIFNIDCART
            WHERE FIFNIDRPT = TB1.FIFNIDRPT AND FIFNCPTD = TB1.IDCONCEPTO AND FIFNIDCIAU = TB1.FIFNIDCIAU AND CFG.FIFNSTATUS = 1) CARTERAS


        FROM(
            SELECT DISTINCT


                PERI.FDCOIDPERI,
                PERI.FICOMES,
                PERI.FICOANIO,
                CONFIGREP.FIFNIDCIAU,
                CONFIGREP.FIFNIDRPT,
                REPORTES.FSFNRPTNM,
                CONFIGREP.FIFNIDAGRP,
                AGRPGENERAL.FSFNIDAGRP,
                CONFIGREP.FIFNIDGRPM,
                AGRPMAESTRO.FSFNAGRPM, CONFIGREP.FIFNIDGRP AS GRUPO, AGRUPADOR.FIFNINDEX AS INDEX,
                AGRUPADOR.FSFNGRPNM AS NOMBREGRUPO,
                CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                CONCEPTOS.FSFNNCPT AS CONCEPTO,
                COALESCE(CONFIG.FIFNIDCART, 0) IDCART,
                COALESCE(CONFIG.FIFNIDCTA, 0) IDCUENTA
            FROM
                PRODCONT.COCPERIO PERI,
                PRODFINA.FNDCFREPCP CONFIGREP
                LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP
                LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM
                LEFT JOIN PRODFINA.FNCGRPCT AGRUPADOR ON CONFIGREP.FIFNIDGRP = AGRUPADOR.FIFNIDGRP
                LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1
                LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT
                LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD AND CONFIG.FIFNSTATUS = 1
            WHERE
                PERI.FICOANIO = " + anio + @" AND  
                CONFIGREP.FIFNIDRPT = 8 AND
                CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"




            ORDER BY CONFIGREP.FIFNIDRPT, CONFIGREP.FIFNIDCPT
        ) TB1
            LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU
                AND CARTERAS.FICCIDCART = TB1.IDCART
                AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                AND CARTERAS.FICCSTATUS = 1
            LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI
                AND CTAS.FICOIDTCCT IN(1)
                AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU
                AND CTAS.FICOIDCTA = TB1.IDCUENTA
                AND CTAS.FICOSTATUS = 1

        WHERE GRUPO = 43 AND IDCONCEPTO IN(754, 755, 762)

        GROUP BY
            TB1.FIFNIDCIAU,
            TB1.FICOANIO,
            TB1.FIFNIDRPT,
            TB1.FSFNRPTNM,
            TB1.FIFNIDAGRP,
            TB1.FSFNIDAGRP,
            TB1.FIFNIDGRPM,
            TB1.FSFNAGRPM,
            TB1.GRUPO, TB1.INDEX,
            TB1.NOMBREGRUPO,
            TB1.IDCONCEPTO,
            TB1.CONCEPTO
        order by TB1.IDCONCEPTO)
        GROUP BY
        GRUPO, INDEX, NOMBREGRUPO

        UNION

        SELECT 45 GRUPO,15 INDEX, 'o - Utilidad Bruta Servicio' NOMBREGRUPO, 757 IDCONCEPTO, 'Utilidad Bruta Servicio' CONCEPTO,  
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 1 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Enero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 2 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Febrero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 3 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Marzo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 4 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Abril,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 5 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Mayo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 6 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Junio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 7 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Julio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 8 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Agosto,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 9 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Septiembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 10 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Octubre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 11 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Noviembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Diciembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES <= 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Total
        , '' CUENTAS, '' CARTERAS
        FROM(
            SELECT FIFNMONTH FICOMES, FIFNVALUE Utilidad_Bruta_Servicio FROM PRODFINA.FNDRESMEN
                WHERE FIFNYEAR = " + anio + @"
                AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                AND FIFNCPTD = 1022
       )
        UNION


        SELECT 46 GRUPO, 16 INDEX, 'p - Utilidad Bruta Refacciones' NOMBREGRUPO, 761 IDCONCEPTO, 'Utilidad Bruta Refacciones' CONCEPTO,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 1 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Enero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 2 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Febrero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 3 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Marzo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 4 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Abril,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 5 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Mayo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 6 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Junio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 7 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Julio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 8 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Agosto,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 9 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Septiembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 10 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Octubre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 11 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Noviembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Diciembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES <= 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Total
        , '' CUENTAS
        , '' CARTERAS
        FROM(
               SELECT FIFNMONTH FICOMES, FIFNVALUE Utilidad_Bruta_Servicio FROM PRODFINA.FNDRESMEN
                WHERE FIFNYEAR = " + anio + @"
                AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                AND FIFNCPTD = 1039
       )


        UNION


        SELECT 47 GRUPO, 17 INDEX, 'q - Utilidad Bruta Hojalateria y Pintura' NOMBREGRUPO, 762 IDCONCEPTO, 'Utilidad Bruta Hojalateria y Pintura' CONCEPTO,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 1 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Enero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 2 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Febrero,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 3 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Marzo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 4 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Abril,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 5 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Mayo,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 6 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Junio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 7 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Julio,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 8 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Agosto,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 9 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Septiembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 10 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Octubre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 11 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Noviembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES = 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Diciembre,
        ROUND(IFNULL(SUM(CASE WHEN FICOMES <= 12 THEN Utilidad_Bruta_Servicio ELSE 0 END), 0) / 1000, 0) Total
        , '' CUENTAS
        , '' CARTERAS
        FROM(
               SELECT FIFNMONTH FICOMES, FIFNVALUE Utilidad_Bruta_Servicio FROM PRODFINA.FNDRESMEN
                WHERE FIFNYEAR = " + anio + @"
                AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                AND FIFNCPTD = 1026
       )

       order by GRUPO, IDCONCEPTO
       )
        GROUP BY
            GRUPO, INDEX,
            NOMBREGRUPO,
            IDCONCEPTO,
            CONCEPTO,
            CUENTAS,
            CARTERAS


        ORDER BY INDEX";

            return dbCnx.GetDataTable(query);
        }

        public void GeneraExcelAG(List<int> idsAgencia, bool esUnaAgencia)
        {
            List<Agencia> agencias;

            if (idsAgencia == null)
            {
                List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1); //Todas las agencias
                List<int> aIdAgencias = new List<int>();

                aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                agencias = Agencia.ListarPorIds(_db, aIdAgencias);
            }
            else
                agencias = Agencia.ListarPorIds(_db, idsAgencia);

            string idAgenciaString = string.Join(", ", agencias.Select(item => item.Id));

            //List<ConceptosContables> conceptosV2 = ConceptosContables.ListarBGV1yV2(_db);
            List<ConceptosContables> conceptosV2 = ConceptosContables.ListarAG_V1yV2(_db);

            DataTable dtBGExcelV1 = GetExcelV1(10);
            DataTable dtBGExcelV1Acum = GetExcelV1Acumulado(10);
            DataTable dtBGExcelV2 = GetExcelV2(10);
            DataTable dtBGExcelV2Acum = GetExcelV2Acumulado(10);

            //DataTable dtBGWebV2Acum = GetBGWebV1();

            int r = 1;
            int c = 1;
            int vExcel = 0;
            int vWeb = 0;
            int intAux = 0;
            DataRow[] drBGExcelV1 = null;
            DataRow[] drBGExcelV1Acum = null;
            DataRow[] drBGExcelV2 = null;
            DataRow[] drBGExcelV2Acum = null;
            DataRow[] drBGWebV1 = null;
            DataRow[] drBGWebV1Acum = null;
            DataRow[] drBGWebV2 = null;
            DataRow[] drBGWebV2Acum = null;

            List<Reporte> listaR = new List<Reporte>();
            List<ReporteGlobal> listaRG = new List<ReporteGlobal>();

            DataTable dtBGWebV1 = new DataTable();
            DataTable dtBGWebV2 = new DataTable();

            var getResTable = GetAGWebV1("28");

            using (DVAExcel.ExcelWriter eW = new DVAExcel.ExcelWriter(ruta))
            {
                foreach (Agencia agencia in agencias.OrderBy(o => o.Siglas))
                {
                    if (esUnaAgencia)
                    {
                        //Metodos a iterar ----------------------------------        
                        dtBGWebV1 = GetAGWebV1(idAgenciaString);

                        dtBGWebV2 = GetAGWebV2(idAgenciaString);
                    }
                    else
                    {
                        dtBGWebV1 = GetAGWebV1(agencia.Id.ToString());

                        dtBGWebV2 = GetAGWebV2(agencia.Id.ToString());
                    }



                    if ((agencia.Id == 590) || (agencia.Id == 583) || (agencia.Id == 301) || (agencia.Id == 583) || (agencia.Id == 563) || (agencia.Id == 593)
                        || (agencia.Id == 592) || (agencia.Id == 594) || (agencia.Id == 100))
                        continue;

                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("[SIGLAS]: " + agencia.Siglas);
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");

                    foreach (ConceptosContables concepto in conceptosV2)
                    {
                        Reporte rep = new Reporte();

                        r++;

                        rep.ID_CONCEPTO = concepto.Id;
                        if (concepto.Id == 691)
                        {
                            var nombreConcepto = concepto.NombreConcepto;
                        }
                        rep.CONCEPTO = concepto.NombreConcepto;

                        int cambioSigo = 1;


                        // drBGExcelV1 = dtBGExcelV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drBGExcelV1 = dtBGExcelV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drBGExcelV1Acum = dtBGExcelV1Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        if (agencia.Id == 27)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 286)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 36)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 12)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 35)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 588)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 32)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 212)
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else
                        {
                            drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        }

                        drBGExcelV2 = dtBGExcelV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drBGExcelV2Acum = dtBGExcelV2Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        if (agencia.Id == 27)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 286)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 36)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 12)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 35)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 588)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 32)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else if (agencia.Id == 212)
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }
                        else
                        {
                            drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                            drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        }

                        if (drBGExcelV1.Length != 0)
                        {
                            vExcel = Convert.ToInt32(drBGExcelV1[0]["VALOR"]);
                            rep.EXCEL_V1 = vExcel;
                            vExcel = Convert.ToInt32(drBGExcelV1Acum[0]["VALOR"]);
                            rep.EXCEL_V1_ACUM = vExcel;
                        }

                        bool esDecimal = false;
                        if (drBGWebV1.Length != 0)
                        {
                            vWeb = 0;

                            if (concepto.Id == 691 || concepto.Id == 692 || concepto.Id == 693 || concepto.Id == 694)
                            {
                                esDecimal = true;
                            }

                            foreach (DataRow dr in drBGWebV1)
                            {
                                switch (mes)
                                {
                                    case 1:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]);
                                        break;
                                    case 2:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) : vWeb += Convert.ToInt32(dr["FEB"]);
                                        break;
                                    case 3:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) : vWeb += Convert.ToInt32(dr["MAR"]);
                                        break;
                                    case 4:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) : vWeb += Convert.ToInt32(dr["ABR"]);
                                        break;
                                    case 5:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) : vWeb += Convert.ToInt32(dr["MAY"]);
                                        break;
                                    case 6:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) : vWeb += Convert.ToInt32(dr["JUN"]);
                                        break;
                                    case 7:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) : vWeb += Convert.ToInt32(dr["JUL"]);
                                        break;
                                    case 8:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) : vWeb += Convert.ToInt32(dr["AGO"]);
                                        break;
                                    case 9:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) : vWeb += Convert.ToInt32(dr["SEP"]);
                                        break;
                                    case 10:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["OCT"]) * 100) : vWeb += Convert.ToInt32(dr["OCT"]);
                                        break;
                                    case 11:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["NOV"]) * 100) : vWeb += Convert.ToInt32(dr["NOV"]);
                                        break;
                                    case 12:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["DIC"]) * 100) : vWeb += Convert.ToInt32(dr["DIC"]);
                                        break;
                                    default:
                                        break;
                                }

                                if (!esDecimal)
                                {
                                    rep.WEB_V1 = vWeb * cambioSigo / 1000;
                                }
                                else
                                {
                                    rep.WEB_V1 = vWeb;
                                }

                            }

                            vWeb = 0;

                            foreach (DataRow dr in drBGWebV1Acum)
                            {
                                if (mes == 1)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]);
                                else if (mes == 2)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]);
                                else if (mes == 3)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]);
                                else if (mes == 4)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]);
                                else if (mes == 5)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]);
                                else if (mes == 6)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]);
                                else if (mes == 7)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]);
                                else if (mes == 8)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]);
                                else if (mes == 9)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]) +
                                        Convert.ToInt32(dr["SEP"]);
                                else if (mes == 10)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["OCT"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]) +
                                        Convert.ToInt32(dr["SEP"]) + Convert.ToInt32(dr["OCT"]);
                                else if (mes == 11)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["OCT"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["NOV"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]) +
                                        Convert.ToInt32(dr["SEP"]) + Convert.ToInt32(dr["OCT"]) +
                                        Convert.ToInt32(dr["NOV"]);
                                else if (mes == 12)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["OCT"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["NOV"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["DIC"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]) +
                                        Convert.ToInt32(dr["SEP"]) + Convert.ToInt32(dr["OCT"]) +
                                        Convert.ToInt32(dr["NOV"]) + Convert.ToInt32(dr["DIC"]);

                                //rep.WEB_V1_ACUM = vWeb * cambioSigo;
                                if (!esDecimal)
                                {
                                    rep.WEB_V1_ACUM = vWeb * cambioSigo / 1000;
                                }
                                else
                                {
                                    rep.WEB_V1_ACUM = vWeb;
                                }
                            }
                        }

                        rep.DIFF_V1 = rep.EXCEL_V1 - rep.WEB_V1;
                        rep.DIFF_V1_ACUM = rep.EXCEL_V1_ACUM - rep.WEB_V1_ACUM;

                        if (drBGExcelV2.Length != 0)
                        {
                            vExcel = Convert.ToInt32(drBGExcelV2[0]["VALOR"]);
                            rep.EXCEL_V2 = vExcel;
                            vExcel = Convert.ToInt32(drBGExcelV2Acum[0]["VALOR"]);
                            rep.EXCEL_V2_ACUM = vExcel;
                        }

                        if (drBGWebV2.Length != 0)
                        {
                            vWeb = 0;

                            foreach (DataRow dr in drBGWebV2)
                            {
                                switch (mes)
                                {
                                    case 1:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]);
                                        break;
                                    case 2:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) : vWeb += Convert.ToInt32(dr["FEB"]);
                                        break;
                                    case 3:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) : vWeb += Convert.ToInt32(dr["MAR"]);
                                        break;
                                    case 4:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) : vWeb += Convert.ToInt32(dr["ABR"]);
                                        break;
                                    case 5:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) : vWeb += Convert.ToInt32(dr["MAY"]);
                                        break;
                                    case 6:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) : vWeb += Convert.ToInt32(dr["JUN"]);
                                        break;
                                    case 7:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) : vWeb += Convert.ToInt32(dr["JUL"]);
                                        break;
                                    case 8:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) : vWeb += Convert.ToInt32(dr["AGO"]);
                                        break;
                                    case 9:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) : vWeb += Convert.ToInt32(dr["SEP"]);
                                        break;
                                    case 10:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["OCT"]) * 100) : vWeb += Convert.ToInt32(dr["OCT"]);
                                        break;
                                    case 11:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["NOV"]) * 100) : vWeb += Convert.ToInt32(dr["NOV"]);
                                        break;
                                    case 12:
                                        vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["DIC"]) * 100) : vWeb += Convert.ToInt32(dr["DIC"]);
                                        break;
                                    default:
                                        break;
                                }

                                //rep.WEB_V2 = vWeb * cambioSigo;
                                if (!esDecimal)
                                {
                                    rep.WEB_V2 = vWeb * cambioSigo / 1000;
                                }
                                else
                                {
                                    rep.WEB_V2 = vWeb;
                                }
                            }

                            vWeb = 0;

                            foreach (DataRow dr in drBGWebV2Acum)
                            {
                                if (mes == 1)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]);
                                else if (mes == 2)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]);
                                else if (mes == 3)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]);
                                else if (mes == 4)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]);
                                else if (mes == 5)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]);
                                else if (mes == 6)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]);
                                else if (mes == 7)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]);
                                else if (mes == 8)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]);
                                else if (mes == 9)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]) +
                                        Convert.ToInt32(dr["SEP"]);
                                else if (mes == 10)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["OCT"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]) +
                                        Convert.ToInt32(dr["SEP"]) + Convert.ToInt32(dr["OCT"]);
                                else if (mes == 11)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["OCT"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["NOV"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]) +
                                        Convert.ToInt32(dr["SEP"]) + Convert.ToInt32(dr["OCT"]) +
                                        Convert.ToInt32(dr["NOV"]);
                                else if (mes == 12)
                                    vWeb += esDecimal ? vWeb += Convert.ToInt32(Convert.ToDecimal(dr["ENE"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["FEB"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["ABR"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["MAY"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUN"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["JUL"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["AGO"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["SEP"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["OCT"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["NOV"]) * 100) + Convert.ToInt32(Convert.ToDecimal(dr["DIC"]) * 100) : vWeb += Convert.ToInt32(dr["ENE"]) + Convert.ToInt32(dr["FEB"]) +
                                        Convert.ToInt32(dr["MAR"]) + Convert.ToInt32(dr["ABR"]) +
                                        Convert.ToInt32(dr["MAY"]) + Convert.ToInt32(dr["JUN"]) +
                                        Convert.ToInt32(dr["JUL"]) + Convert.ToInt32(dr["AGO"]) +
                                        Convert.ToInt32(dr["SEP"]) + Convert.ToInt32(dr["OCT"]) +
                                        Convert.ToInt32(dr["NOV"]) + Convert.ToInt32(dr["DIC"]);

                                //rep.WEB_V2_ACUM = vWeb * cambioSigo;

                                if (!esDecimal)
                                {
                                    rep.WEB_V2_ACUM = vWeb * cambioSigo / 1000;
                                }
                                else
                                {
                                    rep.WEB_V2_ACUM = vWeb;
                                }
                            }
                        }

                        rep.DIFF_V2 = rep.EXCEL_V2 - rep.WEB_V2;
                        rep.DIFF_V2_ACUM = rep.EXCEL_V2_ACUM - rep.WEB_V2_ACUM;

                        Console.WriteLine("[CONCEPTO]: " + concepto.Id + "_" + concepto.NombreConcepto +
                            " [EXCEL_V1]: " + rep.EXCEL_V1 +
                            " [WEB_V1]: " + rep.WEB_V1 +
                            " [DIFF_V1]: " + rep.DIFF_V1 +
                            " [EXCEL_V1_ACUMULADO]: " + rep.EXCEL_V1_ACUM +
                            " [WEB_V1_ACUMULADO]: " + rep.WEB_V1_ACUM +
                            " [DIFF_V1_ACUMULADO]: " + rep.DIFF_V1_ACUM +
                            " [EXCEL_V2]: " + rep.EXCEL_V2 +
                            " [WEB_V2]: " + rep.WEB_V2 +
                            " [DIFF_V2]: " + rep.DIFF_V2 +
                            " [EXCEL_V2_ACUMULADO]: " + rep.EXCEL_V2_ACUM +
                            " [WEB_V2_ACUMULADO]: " + rep.WEB_V2_ACUM +
                            " [DIFF_V2_ACUMULADO]: " + rep.DIFF_V2_ACUM);

                        listaR.Add(rep);
                    }

                    var dtToExcel = listaR.ToDataTable();
                    eW.WriteDataTable(listaR.ToDataTable(), agencia.Siglas);
                    eW.Dispose();

                    listaR = new List<Reporte>();
                }
            }
        }

        public DataTable GetAGWebV1(string Id_Agencia)
        {
            string query = @"select 
    1 INDX
   -- @AG IDAGENCIA
    ,vp.IDNOMBREGRUPO
    ,vp.NOMBREGRUPO
    ,vp.IdConcepto
    ,vp.CONCEPTO
    ,ifnull(vp.ENE,0) ENE
    ,ifnull(vp.FEB,0) FEB
    ,ifnull(vp.MAR,0) MAR
    ,ifnull(vp.ABR,0) ABR
    ,ifnull(vp.MAY,0) MAY
    ,ifnull(vp.JUN,0) JUN
    ,ifnull(vp.JUL,0) JUL
    ,ifnull(vp.AGO,0) AGO
    ,ifnull(vp.SEP,0) SEP
    ,ifnull(vp.Oct,0) OCT
    ,ifnull(vp.NOV,0) NOV
    ,ifnull(vp.DIC,0) DIC
    ,ifnull(vp.Total,0) TOTAL


    from (


     Select
 
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,1 IDCONCEPTOORDEN
    , 668 IdConcepto 
    ,'VENTAS AGENCIA MES' CONCEPTO
    ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
    ,SUM(CASE WHEN FIFNMONTH = 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
    ,SUM(CASE WHEN FIFNMONTH = 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
    ,SUM(CASE WHEN FIFNMONTH = 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
    ,SUM(CASE WHEN FIFNMONTH = 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
    ,SUM(CASE WHEN FIFNMONTH = 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
    ,SUM(CASE WHEN FIFNMONTH = 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
    ,SUM(CASE WHEN FIFNMONTH = 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
    ,SUM(CASE WHEN FIFNMONTH = 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
    ,SUM(CASE WHEN FIFNMONTH = 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
    ,SUM(CASE WHEN FIFNMONTH = 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
    ,SUM(CASE WHEN FIFNMONTH = 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
    ,SUM(COALESCE(FIFNVALUE,0)) TOTAL

    FROM PRODFINA.FNDRESMEN
    WHERE 
    FIFNYEAR = " + anio + @"
    AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
    AND FIFNCPTD = 1000
    AND FIFNSTATUS = 1
    

    UNION ALL

    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,2 IDCONCEPTOORDEN
    ,669 IdConcepto
    ,'VENTAS AGENCIA ACUMULADO' CONCEPTO
    ,SUM(CASE WHEN FIFNMONTH = 1 THEN    COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
    ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
    ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
    ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
    ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
    ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
    ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
    ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
    ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
    ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
    ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
    ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
    ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
    FROM PRODFINA.FNDRESMEN
    WHERE 
    FIFNYEAR = " + anio + @"
    AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
    AND FIFNCPTD = 1000
    AND FIFNSTATUS = 1

    union ALL

    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,3 IDCONCEPTOORDEN
    ,670 IdConcepto
    ,'VENTAS FIDEICOMISO MES' CONCEPTO,
    SUM(CASE WHEN FICOMES = 1 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) ENE,
    SUM(CASE WHEN FICOMES = 2 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) FEB,
    SUM(CASE WHEN FICOMES = 3 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) MAR,
    SUM(CASE WHEN FICOMES = 4 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) ABR,
    SUM(CASE WHEN FICOMES = 5 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) MAY,
    SUM(CASE WHEN FICOMES = 6 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) JUN,
    SUM(CASE WHEN FICOMES = 7 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) JUL,
    SUM(CASE WHEN FICOMES = 8 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) AGO,
    SUM(CASE WHEN FICOMES = 9 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) SEP,
    SUM(CASE WHEN FICOMES = 10 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) OCT,
    SUM(CASE WHEN FICOMES = 11 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) NOV,
    SUM(CASE WHEN FICOMES = 12 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) DIC
    ,SUM(CASE WHEN FICOMES <= 12 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) TOTAL

    FROM (SELECT PERI.FICOMES,
           
           (sum(case when CTADESC.FSCOCUENTA IN ('400000010005') then DECIMAL(((CTAS.FDCOTOTCAR) - (CTAS.FDCOTOTABO)),20,0) else 0 end)                                                                                                                                                                          
           ) INGRESOS_AUTOS_NUEVOS
    
           FROM  PRODCONT.COESLDCT CTAS
           left join  PRODCONT.COCATCTS CTADESC ON CTAS.FICOIDCTA = CTADESC.FICOIDCTA 
           LEFT JOIN  PRODCONT.COCPERIO PERI on CTAS.FDCOIDPERI = PERI.FDCOIDPERI  
           LEFT JOIN (SELECT  CIA.FIGEIDCIAU,CIA.FSGERAZSOC,CIA.FSGESIGCIA,CIA.FSGECALLE,CIA.FIGEIDCLAS,CIA.FIGEIDDIVI,CIA.FIGEIDMARC,MARCA.FNMARCCLA,MARCA.FSMARCDES,MARCA.FNMARCASO,MARCA.FIANSTATU,MARCA.FIANIDCLAS
                      FROM  PRODGRAL.GECCIAUN CIA 
                      LEFT JOIN  PRODAUT.ANCMARCA MARCA  
                      ON CIA.FIGEIDCLAS = MARCA.FIANIDCLAS 
                      and CIA.FIGEIDMARC = MARCA.FNMARCCLA) MARCAS ON MARCAS.FIGEIDCIAU = CTAS.FICOIDCIAU 
            WHERE  
            CTAS.FICOIDTCCT = 1 
            AND CTAS.FICOIDCIAU IN(" + Id_Agencia + ')' + @"
            AND CTAS.FICOSTATUS = 1
            AND FICOANIO = " + anio + @"
            GROUP BY
            PERI.FICOMES
            ORDER BY FICOMES ASC)
    UNION ALL

    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,4 IDCONCEPTOORDEN
    ,671 IdConcepto
    ,'VENTAS FIDEICOMISO ACUMULADO' CONCEPTO,
    SUM(CASE WHEN FICOMES = 1 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) ENE,
    SUM(CASE WHEN FICOMES <= 2 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) FEB,
    SUM(CASE WHEN FICOMES <= 3 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) MAR,
    SUM(CASE WHEN FICOMES <= 4 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) ABR,
    SUM(CASE WHEN FICOMES <= 5 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) MAY,
    SUM(CASE WHEN FICOMES <= 6 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) JUN,
    SUM(CASE WHEN FICOMES <= 7 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) JUL,
    SUM(CASE WHEN FICOMES <= 8 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) AGO,
    SUM(CASE WHEN FICOMES <= 9 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) SEP,
    SUM(CASE WHEN FICOMES <= 10 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) OCT,
    SUM(CASE WHEN FICOMES <= 11 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) NOV,
    SUM(CASE WHEN FICOMES <= 12 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) DIC
    ,SUM(CASE WHEN FICOMES <= 12 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) TOTAL

    FROM (SELECT PERI.FICOMES,
           
           (sum(case when CTADESC.FSCOCUENTA IN ('400000010005') then DECIMAL(((CTAS.FDCOTOTCAR) - (CTAS.FDCOTOTABO)),20,0) else 0 end)                                                                                                                                                                          
           ) INGRESOS_AUTOS_NUEVOS
    
           FROM  PRODCONT.COESLDCT CTAS
           left join  PRODCONT.COCATCTS CTADESC ON CTAS.FICOIDCTA = CTADESC.FICOIDCTA 
           LEFT JOIN  PRODCONT.COCPERIO PERI on CTAS.FDCOIDPERI = PERI.FDCOIDPERI  
           LEFT JOIN (SELECT  CIA.FIGEIDCIAU,CIA.FSGERAZSOC,CIA.FSGESIGCIA,CIA.FSGECALLE,CIA.FIGEIDCLAS,CIA.FIGEIDDIVI,CIA.FIGEIDMARC,MARCA.FNMARCCLA,MARCA.FSMARCDES,MARCA.FNMARCASO,MARCA.FIANSTATU,MARCA.FIANIDCLAS
                      FROM  PRODGRAL.GECCIAUN CIA 
                      LEFT JOIN  PRODAUT.ANCMARCA MARCA  
                      ON CIA.FIGEIDCLAS = MARCA.FIANIDCLAS 
                      and CIA.FIGEIDMARC = MARCA.FNMARCCLA) MARCAS ON MARCAS.FIGEIDCIAU = CTAS.FICOIDCIAU 
            WHERE  
            CTAS.FICOIDTCCT = 1 
            AND CTAS.FICOIDCIAU IN(" + Id_Agencia + ')' + @"
            AND CTAS.FICOSTATUS = 1
            AND FICOANIO = " + anio + @"
            GROUP BY
            PERI.FICOMES
            ORDER BY FICOMES ASC)
    UNION ALL 

    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,5 IDCONCEPTOORDEN
    ,672 IdConcepto
    ,'TOTAL VENTAS CON FIDEICOMISO ACUMULADO' CONCEPTO
    ,SUM(CASE WHEN T1.FIFNMONTH  =  1  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH =  1  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) ENE
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 2  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 2  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) FEB
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 3  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 3  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) MAR
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 4  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 4  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) ABR
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 5  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 5  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) MAY
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 6  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 6  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) JUN
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 7  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 7  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) JUL
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 8  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 8  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) AGO
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 9  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 9  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) SEP
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 10 AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 10 THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) OCT
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 11 AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 11 THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) NOV
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 12 AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 12 THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) DIC
    ,SUM(CASE WHEN T1.FIFNCPTD = 1000 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM(COALESCE(decimal(T2.FIFNVTAAUF,20,0),0)) TOTAL

    FROM PRODFINA.FNDRESMEN T1 LEFT JOIN 
    PRODFINA.FNDRPTMST T2 ON T1.FIFNYEAR = T2.FIFNYEAR 
    AND T1.FIFNMONTH = T2.FIFNMONTH AND T1.FIFNIDCIAU = T2.FIFNIDCIAU

    WHERE 
    T1.FIFNYEAR = " + anio + @"
    AND T1.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
    AND FIFNCPTD = 1000
    AND T1.FIFNSTATUS = 1
    union ALL
    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,6 IDCONCEPTOORDEN
    ,673 IdConcepto
    ,'UNIDADES NUEVAS VENDIDAS AGENCIA MES' CONCEPTO
    ,SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 ENE
    ,SUM(CASE WHEN FIFNMONTH = 2 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 FEB
    ,SUM(CASE WHEN FIFNMONTH = 3 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 MAR
    ,SUM(CASE WHEN FIFNMONTH = 4 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 ABR
    ,SUM(CASE WHEN FIFNMONTH = 5 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 MAY
    ,SUM(CASE WHEN FIFNMONTH = 6 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 JUN
    ,SUM(CASE WHEN FIFNMONTH = 7 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 JUL
    ,SUM(CASE WHEN FIFNMONTH = 8 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 AGO
    ,SUM(CASE WHEN FIFNMONTH = 9 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 SEP
    ,SUM(CASE WHEN FIFNMONTH = 10 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 Oct
    ,SUM(CASE WHEN FIFNMONTH = 11 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 NOV
    ,SUM(CASE WHEN FIFNMONTH = 12 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 DIC
    ,SUM(decimal(FIFNVALUE,20,0))*1000 TOTAL

    FROM PRODFINA.FNDRESMEN
    WHERE 
    FIFNYEAR = " + anio + @"
    AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
    AND FIFNCPTD = 1001
    AND FIFNSTATUS = 1
    union ALL

    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,7 IDCONCEPTOORDEN
    ,674 IdConcepto
    ,'UNIDADES NUEVAS VENDIDAS AUTOFIN MES' CONCEPTO
    ,SUM(CASE WHEN MES = 1 THEN   COALESCE(UND,0) ELSE 0 END ) ENE
    ,SUM(CASE WHEN MES = 2 THEN  COALESCE(UND,0) ELSE 0 END ) FEB
    ,SUM(CASE WHEN MES = 3 THEN  COALESCE(UND,0) ELSE 0 END ) MAR
    ,SUM(CASE WHEN MES = 4 THEN  COALESCE(UND,0) ELSE 0 END ) ABR
    ,SUM(CASE WHEN MES = 5 THEN  COALESCE(UND,0) ELSE 0 END ) MAY
    ,SUM(CASE WHEN MES = 6 THEN  COALESCE(UND,0) ELSE 0 END ) JUN
    ,SUM(CASE WHEN MES = 7 THEN  COALESCE(UND,0) ELSE 0 END ) JUL
    ,SUM(CASE WHEN MES = 8 THEN  COALESCE(UND,0) ELSE 0 END ) AGO
    ,SUM(CASE WHEN MES = 9 THEN  COALESCE(UND,0) ELSE 0 END ) SEP
    ,SUM(CASE WHEN MES = 10 THEN COALESCE(UND,0)  ELSE 0 END ) Oct
    ,SUM(CASE WHEN MES = 11 THEN COALESCE(UND,0)  ELSE 0 END ) NOV
    ,SUM(CASE WHEN MES = 12 THEN COALESCE(UND,0)  ELSE 0 END ) DIC
    ,SUM(COALESCE(UND,0)) TOTAL
    FROM (
    SELECT * FROM(
        SELECT  COALESCE(FICAIDCIAU,FIBIIDCIAU) IDCIAU, COALESCE(AÑO,YEAR(FFBIFECHA)) AÑO, COALESCE(MES,MONTH(FFBIFECHA)) MES, COALESCE(UNIDADES,0) UND
        FROM   
        (
            SELECT  FICAIDCIAU, AÑO, MES, SUM(UNIDADES) UNIDADES , sum(FDCASUBTOT) VENTA, SUM(FDCATOTAL) VENTA_F
            FROM   (
                SELECT  FA.FICAIDCIAU,YEAR(FFCAFECHA) AÑO, MONTH(FFCAFECHA) MES, sum(FA.FDCASUBTOT) FDCASUBTOT ,sum(FA.FDCATOTAL) FDCATOTAL ,  COUNT(*) UNIDADES
                FROM PRODCAJA.CAEFACTU FA
                INNER JOIN        PRODCAJA.CAEFACAN FN
                ON         FA.FICAIDCIAU=FN.FICAIDCIAU AND FA.FICAIDFACT=FN.FICAIDFACT
                LEFT JOIN           PRODCAJA.CAENOTCR NC
                ON         FA.FICAIDCIAU=NC.FICAIDCIAU AND FA.FICAIDFACT=NC.FICAIDFACT
                AND       NC.FICAIDTINC=3 AND FICAIDESTA=4
                WHERE FA.FICAIDTIFA=3 
                                AND (NC.FICAIDFACT IS NOT NULL OR FA.FICASTATUS=1) 
                                AND YEAR(FFCAFECHA) IN (" + anio + @")                                                                --AND MONTH(FFCAFECHA) IN (8)
                                AND FSCAMARCA NOT IN ('AIMA','BAJAJ') 
                                AND FA.FICAIDCIAU IN(" + Id_Agencia + ')' + @"
                                AND FICAIDTIVT IN (3)                                                       

                GROUP BY  FA.FICAIDCIAU, YEAR(FFCAFECHA) ,MONTH(FFCAFECHA)
                UNION
                SELECT  FA.FICAIDCIAU, YEAR(FFCAFECHAE) AÑO, MONTH(FFCAFECHAE) MES,sum(FA.FDCASUBTOT)*-1 FDCASUBTOT, sum(FA.FDCATOTAL)*-1 FDCATOTAL, COUNT(*)*-1 UNIDADES
                FROM   PRODCAJA.CAEFACTU FA
                INNER JOIN        PRODCAJA.CAEFACAN FN
                ON         FA.FICAIDCIAU=FN.FICAIDCIAU AND FA.FICAIDFACT=FN.FICAIDFACT
                INNER JOIN        PRODCAJA.CAENOTCR NC
                ON         FA.FICAIDCIAU=NC.FICAIDCIAU AND FA.FICAIDFACT=NC.FICAIDFACT
                AND       NC.FICAIDTINC=3 AND FICAIDESTA=4
                WHERE NC.FICASTATUS=1
                                AND FA.FICAIDTIFA=3 
                                AND YEAR(FFCAFECHAE) IN (" + anio + @")
                                AND FSCAMARCA NOT IN ('AIMA','BAJAJ') 
                                AND FA.FICAIDCIAU IN(" + Id_Agencia + ')' + @"
                                AND FICAIDTIVT IN (3)
                GROUP BY          FA.FICAIDCIAU, YEAR(FFCAFECHAE) ,MONTH(FFCAFECHAE)
            ) X
            GROUP BY          FICAIDCIAU, AÑO, MES
        )S
        FULL JOIN (
                    SELECT * FROM PRODBUIN.BIDINDAG WHERE FIBIIDINDI=52 AND FIBIIDTIPO=2 AND FIBIIDAREA=1 AND YEAR(FFBIFECHA) IN (" + anio + @")
                ) IND
        ON  FICAIDCIAU=FIBIIDCIAU AND AÑO=YEAR(FFBIFECHA) AND MES=MONTH(FFBIFECHA)
        WHERE 
        FICAIDCIAU IN(" + Id_Agencia + ')' + @"
        ) WHERE AÑO = " + anio + @"
    )

    UNION ALL 

    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,8 IDCONCEPTOORDEN
    ,675 IdConcepto
    ,'UNIDADES NUEVAS VENDIDAS TAXIS BAM  MES' CONCEPTO
    ,SUM(0) ENE
    ,SUM(0) FEB
    ,SUM(0) MAR
    ,SUM(0) ABR
    ,SUM(0) MAY
    ,SUM(0) JUN
    ,SUM(0) JUL
    ,SUM(0) AGO
    ,SUM(0) SEP
    ,SUM(0)  Oct
    ,SUM(0)  NOV
    ,SUM(0) DIC
    ,SUM(0) TOTAL
    FROM PRODFINA.FNDRPTMST
    WHERE 
    FIFNYEAR = " + anio + @"
    AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"


    UNION ALL
 SELECT 1 IDNOMBREGRUPO
        ,'AG' NOMBREGRUPO
        ,9 IDCONCEPTOORDEN
        ,676 IdConcepto
        ,'TOTAL UNIDADES NUEVAS VENDIDAS MES' CONCEPTO
        ,SUM(ENE)*1000 ENE
        ,SUM(FEB)*1000 FEB
        ,SUM(MAR)*1000 MAR
        ,SUM(ABR)*1000 ABR
        ,SUM(MAY)*1000 MAY
        ,SUM(JUN)*1000 JUN
        ,SUM(JUL)*1000 JUL
        ,SUM(AGO)*1000 AGO
        ,SUM(SEP)*1000 SEP
        ,SUM(OCT)*1000 OCT
        ,SUM(NOV)*1000 NOV
        ,SUM(DIC)*1000 DIC
        ,SUM(TOTAL)*1000 TOTAL
        FROM (
            SELECT 
 SUM(ENE) ENE
, SUM(FEB) FEB
, SUM(MAR) MAR
, SUM(ABR) ABR
, SUM(MAY) MAY
, SUM(JUN) JUN
, SUM(JUL) JUL
, SUM(AGO) AGO
, SUM(SEP) SEP
, SUM(OCT) OCT
, SUM(NOV) NOV
, SUM(DIC) DIC
, SUM(TOTAL) TOTAL
FROM(
SELECT 
                    SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ENE
                    ,SUM(CASE WHEN FIFNMONTH = 2 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) FEB
                    ,SUM(CASE WHEN FIFNMONTH = 3 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAR
                    ,SUM(CASE WHEN FIFNMONTH = 4 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ABR
                    ,SUM(CASE WHEN FIFNMONTH = 5 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAY
                    ,SUM(CASE WHEN FIFNMONTH = 6 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUN
                    ,SUM(CASE WHEN FIFNMONTH = 7 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUL
                    ,SUM(CASE WHEN FIFNMONTH = 8 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) AGO
                    ,SUM(CASE WHEN FIFNMONTH = 9 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) SEP
                    ,SUM(CASE WHEN FIFNMONTH = 10 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) OCT
                    ,SUM(CASE WHEN FIFNMONTH = 11 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) NOV
                    ,SUM(CASE WHEN FIFNMONTH = 12 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) DIC
                    ,SUM(decimal(FIFNVALUE,20,0)) TOTAL
                   FROM PRODFINA.FNDVIHIST WHERE FIFNIDCIAU IN(" + Id_Agencia + ')' + @" AND FIFNYEAR = " + anio + @"  AND FIFNCAMPO IN (1,4) AND FIFNSTATUS = 1 GROUP BY FIFNMONTH
         ) )

    UNION ALL



    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,10 IDCONCEPTOORDEN
    ,677 IdConcepto
    ,'TOTAL UNIDADES NUEVAS VENDIDAS ACUMULADO' CONCEPTO
    ,SUM(ENE)*1000 ENE
        ,SUM(FEB)*1000 FEB
        ,SUM(MAR)*1000 MAR
        ,SUM(ABR)*1000 ABR
        ,SUM(MAY)*1000 MAY
        ,SUM(JUN)*1000 JUN
        ,SUM(JUL)*1000 JUL
        ,SUM(AGO)*1000 AGO
        ,SUM(SEP)*1000 SEP
        ,SUM(OCT)*1000 OCT
        ,SUM(NOV)*1000 NOV
        ,SUM(DIC)*1000 DIC
        ,SUM(TOTAL)*1000 TOTAL
        FROM (
            SELECT 
 SUM(ENE) ENE
, SUM(FEB) FEB
, SUM(MAR) MAR
, SUM(ABR) ABR
, SUM(MAY) MAY
, SUM(JUN) JUN
, SUM(JUL) JUL
, SUM(AGO) AGO
, SUM(SEP) SEP
, SUM(OCT) OCT
, SUM(NOV) NOV
, SUM(DIC) DIC
, SUM(TOTAL) TOTAL
FROM(
SELECT 
                    SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ENE
                    ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) FEB
                    ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAR
                    ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ABR
                    ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAY
                    ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUN
                    ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUL
                    ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) AGO
                    ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) SEP
                    ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) OCT
                    ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) NOV
                    ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) DIC
                    ,SUM(decimal(FIFNVALUE,20,0)) TOTAL
                   FROM PRODFINA.FNDVIHIST WHERE FIFNIDCIAU IN(" + Id_Agencia + ')' + @" AND FIFNYEAR = " + anio + @"  AND FIFNCAMPO IN (1,4) AND FIFNSTATUS = 1 GROUP BY FIFNMONTH
         ) )

    UNION ALL

    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,11 IDCONCEPTOORDEN
    ,678 IdConcepto
    ,'PROMEDIO UTILIDAD BRUTA POR UNIDAD NUEVA VENDIDA ACUMULADO' CONCEPTO
    ,SUM(CASE WHEN T1.FIFNMONTH  =  1  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH  = 1  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) ENE
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 2  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 2  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) FEB
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 3  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 3  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) MAR
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 4  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 4  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) ABR
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 5  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 5  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) MAY
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 6  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 6  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) JUN
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 7  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 7  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) JUL
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 8  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 8  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END)  AGO
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 9  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 9  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END)  SEP
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 10 AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 10 AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) Oct
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 11 AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 11 AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) NOV
    ,SUM(CASE WHEN T1.FIFNMONTH  <= 12 AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 12 AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) DIC
    ,SUM(CASE WHEN T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNCPTD =  1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END ) TOTAL

 
                        FROM PRODFINA.FNDRESMEN T1 LEFT JOIN  
                        PRODFINA.FNDRPTMST T2 ON T1.FIFNYEAR = T2.FIFNYEAR
                        AND T1.FIFNMONTH = T2.FIFNMONTH AND T1.FIFNIDCIAU = T2.FIFNIDCIAU 
 
                        WHERE  
                        T1.FIFNYEAR = " + anio + @" 
                        AND T1.FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
                        AND T1.FIFNCPTD IN  (1001,1004) 
 AND T1.FIFNSTATUS = 1
    union ALL

    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,12 IDCONCEPTOORDEN
    ,679 IdConcepto
    ,'UTILIDAD BRUTA PROMEDIO ACUMULADO' CONCEPTO
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH = 1    AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH = 1    AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  ENE  
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 2   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 2   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  FEB  
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 3   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 3   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  MAR
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 4   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 4   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  ABR
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 5   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 5   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  MAY
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 6   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 6   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  JUN
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 7   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 7   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  JUL
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 8   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 8   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  AGO
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 9   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 9   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  SEP
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 10  AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 10  AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  Oct
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 11  AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 11  AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  NOV
    ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 12  AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 12  AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END ),20)  DIC
    ,decimal(SUM(CASE WHEN T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END ),20)  TOTAL



                        FROM PRODFINA.FNDRESMEN T1
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD IN  (1004,1002)
 AND T1.FIFNSTATUS = 1

    union ALL
    Select
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,13 IDCONCEPTOORDEN
    ,680 IdConcepto
    ,'GASTO CORPORATIVO DEL MES' CONCEPTO
    ,SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ENE
    ,SUM(CASE WHEN FIFNMONTH = 2 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) FEB
    ,SUM(CASE WHEN FIFNMONTH = 3 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAR
    ,SUM(CASE WHEN FIFNMONTH = 4 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ABR
    ,SUM(CASE WHEN FIFNMONTH = 5 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAY
    ,SUM(CASE WHEN FIFNMONTH = 6 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUN
    ,SUM(CASE WHEN FIFNMONTH = 7 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUL
    ,SUM(CASE WHEN FIFNMONTH = 8 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) AGO
    ,SUM(CASE WHEN FIFNMONTH = 9 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) SEP
    ,SUM(CASE WHEN FIFNMONTH = 10 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) Oct
    ,SUM(CASE WHEN FIFNMONTH = 11 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) NOV
    ,SUM(CASE WHEN FIFNMONTH = 12 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) DIC
    ,SUM(decimal(FIFNVALUE,20,0)) TOTAL

                        FROM PRODFINA.FNDRESMEN
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD = 1050
 AND FIFNSTATUS = 1
                        union ALL
                        Select
                        1 IDNOMBREGRUPO
                        ,'AG' NOMBREGRUPO
                        ,14 IDCONCEPTOORDEN
                        ,681 IdConcepto
                        ,'UTILIDAD ANTES DEL FIDEICOMISO MES' CONCEPTO
                        ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                        ,SUM(CASE WHEN FIFNMONTH = 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                        ,SUM(CASE WHEN FIFNMONTH = 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                        ,SUM(CASE WHEN FIFNMONTH = 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                        ,SUM(CASE WHEN FIFNMONTH = 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                        ,SUM(CASE WHEN FIFNMONTH = 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                        ,SUM(CASE WHEN FIFNMONTH = 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                        ,SUM(CASE WHEN FIFNMONTH = 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                        ,SUM(CASE WHEN FIFNMONTH = 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                        ,SUM(CASE WHEN FIFNMONTH = 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                        ,SUM(CASE WHEN FIFNMONTH = 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                        ,SUM(CASE WHEN FIFNMONTH = 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                        ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                        FROM PRODFINA.FNDRESMEN
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD = 1051
 AND FIFNSTATUS = 1
    union ALL
    
    SELECT 
    TB1.IDNOMBREGRUPO
    ,TB1.NOMBREGRUPO
    ,TB1.IDCONCEPTOORDEN
    ,TB1.IdConcepto
    ,TB1.CONCEPTO
    ,IFNULL(TB1.ENE,0) ENE
    ,IFNULL(TB1.FEB,0) FEB
    ,IFNULL(TB1.MAR,0) MAR
    ,IFNULL(TB1.ABR,0) ABR
    ,IFNULL(TB1.MAY,0) MAY
    ,IFNULL(TB1.JUN,0) JUN
    ,IFNULL(TB1.JUL,0) JUL
    ,IFNULL(TB1.AGO,0) AGO
    ,IFNULL(TB1.SEP,0) SEP
    ,IFNULL(TB1.OCT,0) OCT
    ,IFNULL(TB1.NOV,0) NOV
    ,IFNULL(TB1.DIC,0) DIC
    ,IFNULL(TB1.TOTAL,0) TOTAL




    FROM (


                        Select
                                                    1 IDNOMBREGRUPO
                                                    ,'AG' NOMBREGRUPO
                                                    ,15 IDCONCEPTOORDEN
                                                    ,682 IdConcepto
                                                    ,'UTILIDAD FIDEICOMISO MES' CONCEPTO
                                                    ,SUM(CASE WHEN FIFNMONTH = 1  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                                                    ,SUM(CASE WHEN FIFNMONTH = 2  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                                                    ,SUM(CASE WHEN FIFNMONTH = 3  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                                                    ,SUM(CASE WHEN FIFNMONTH = 4  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                                                    ,SUM(CASE WHEN FIFNMONTH = 5  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                                                    ,SUM(CASE WHEN FIFNMONTH = 6  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                                                    ,SUM(CASE WHEN FIFNMONTH = 7  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                                                    ,SUM(CASE WHEN FIFNMONTH = 8  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                                                    ,SUM(CASE WHEN FIFNMONTH = 9  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                                                    ,SUM(CASE WHEN FIFNMONTH = 10  THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                                                    ,SUM(CASE WHEN FIFNMONTH = 11  THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                                                    ,SUM(CASE WHEN FIFNMONTH = 12  THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                                                    ,SUM(COALESCE(IFNULL(FIFNVALUE,0),0)) TOTAL
                                                    FROM PRODFINA.FNDRESMEN
                                                    WHERE 
                                                    FIFNYEAR = " + anio + @"
                                                    AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                                    AND FIFNCPTD = 1052

    ) TB1

    UNION ALL 

    SELECT 
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,16 IDCONCEPTOORDEN
    ,683 IdConcepto
    ,'PARTIDAS EXTRAORDINARIAS MES' CONCEPTO
    ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 ENE
    ,SUM(CASE WHEN FIFNMONTH = 2 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 FEB
    ,SUM(CASE WHEN FIFNMONTH = 3 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 MAR
    ,SUM(CASE WHEN FIFNMONTH = 4 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 ABR
    ,SUM(CASE WHEN FIFNMONTH = 5 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 MAY
    ,SUM(CASE WHEN FIFNMONTH = 6 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 JUN
    ,SUM(CASE WHEN FIFNMONTH = 7 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 JUL
    ,SUM(CASE WHEN FIFNMONTH = 8 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 AGO
    ,SUM(CASE WHEN FIFNMONTH = 9 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 SEP
    ,SUM(CASE WHEN FIFNMONTH = 10 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 Oct
    ,SUM(CASE WHEN FIFNMONTH = 11 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 NOV
    ,SUM(CASE WHEN FIFNMONTH = 12 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 DIC
    ,SUM(COALESCE(decimal(FIFNVALMOV,20),0))*1000 TOTAL

    FROM 

    PRODFINA.FNCPXRESCP 

    WHERE FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
    AND FIFNYEAR = " + anio + @"
    AND FIFNSTSEG = 1
    AND FIFNSTATUS = 1

                        union ALL
                        Select
                        1 IDNOMBREGRUPO
                        ,'AG' NOMBREGRUPO
                        ,17 IDCONCEPTOORDEN
                        ,684 IdConcepto
                        ,'UTILIDAD FINAL ANTES DE IMPUESTOS MES' CONCEPTO
                        ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                        ,SUM(CASE WHEN FIFNMONTH = 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                        ,SUM(CASE WHEN FIFNMONTH = 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                        ,SUM(CASE WHEN FIFNMONTH = 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                        ,SUM(CASE WHEN FIFNMONTH = 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                        ,SUM(CASE WHEN FIFNMONTH = 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                        ,SUM(CASE WHEN FIFNMONTH = 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                        ,SUM(CASE WHEN FIFNMONTH = 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                        ,SUM(CASE WHEN FIFNMONTH = 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                        ,SUM(CASE WHEN FIFNMONTH = 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                        ,SUM(CASE WHEN FIFNMONTH = 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                        ,SUM(CASE WHEN FIFNMONTH = 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                        ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                        FROM PRODFINA.FNDRESMEN
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD = 1055


                        union ALL
                        Select
                        1 IDNOMBREGRUPO
                        ,'AG' NOMBREGRUPO
                        ,18 IDCONCEPTOORDEN
                        ,685 IdConcepto
                        ,'GASTO CORPORATIVO ACUMULADO' CONCEPTO
                        ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                        ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                        ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                        ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                        ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                        ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                        ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                        ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                        ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                        ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                        ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                        ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                        ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                        FROM PRODFINA.FNDRESMEN
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD = 1050


    union ALL

    SELECT 
    TB1.IDNOMBREGRUPO
    ,TB1.NOMBREGRUPO
    ,TB1.IDCONCEPTOORDEN
    ,TB1.IdConcepto
    ,TB1.CONCEPTO
    ,IFNULL(TB1.ENE,0) ENE
    ,IFNULL(TB1.FEB,0) FEB
    ,IFNULL(TB1.MAR,0) MAR
    ,IFNULL(TB1.ABR,0) ABR
    ,IFNULL(TB1.MAY,0) MAY
    ,IFNULL(TB1.JUN,0) JUN
    ,IFNULL(TB1.JUL,0) JUL
    ,IFNULL(TB1.AGO,0) AGO
    ,IFNULL(TB1.SEP,0) SEP
    ,IFNULL(TB1.OCT,0) OCT
    ,IFNULL(TB1.NOV,0) NOV
    ,IFNULL(TB1.DIC,0) DIC
    ,IFNULL(TB1.TOTAL,0) TOTAL




    FROM (



                        Select
                        1 IDNOMBREGRUPO
                        ,'AG' NOMBREGRUPO
                        ,19 IDCONCEPTOORDEN
                        ,686 IdConcepto
                        ,'UTILIDAD ANTES DE FIDEICOMISO ACUMULADO' CONCEPTO
                        ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                        ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                        ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                        ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                        ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                        ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                        ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                        ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                        ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                        ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                        ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                        ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                        ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                        FROM PRODFINA.FNDRESMEN
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD = 1051

    ) TB1

    union ALL

    SELECT 
    TB1.IDNOMBREGRUPO
    ,TB1.NOMBREGRUPO
    ,TB1.IDCONCEPTOORDEN
    ,TB1.IdConcepto
    ,TB1.CONCEPTO
    ,IFNULL(TB1.ENE,0) ENE
    ,IFNULL(TB1.FEB,0) FEB
    ,IFNULL(TB1.MAR,0) MAR
    ,IFNULL(TB1.ABR,0) ABR
    ,IFNULL(TB1.MAY,0) MAY
    ,IFNULL(TB1.JUN,0) JUN
    ,IFNULL(TB1.JUL,0) JUL
    ,IFNULL(TB1.AGO,0) AGO
    ,IFNULL(TB1.SEP,0) SEP
    ,IFNULL(TB1.OCT,0) OCT
    ,IFNULL(TB1.NOV,0) NOV
    ,IFNULL(TB1.DIC,0) DIC
    ,IFNULL(TB1.TOTAL,0) TOTAL

                        FROM (
                        Select
                        1 IDNOMBREGRUPO
                        ,'AG' NOMBREGRUPO
                        ,20 IDCONCEPTOORDEN
                        ,687 IdConcepto
                        ,'UTILIDAD FIDEICOMISO ACUMULADA' CONCEPTO
                        ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                        ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                        ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                        ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                        ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                        ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                        ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                        ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                        ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                        ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                        ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                        ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                        ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                        FROM PRODFINA.FNDRESMEN
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD = 1052
                        ) TB1

    UNION ALL 

    SELECT 
    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,21 IDCONCEPTOORDEN
    ,688 IdConcepto
    ,'PARTIDAS EXTRAORDINARIAS ACUMULADO' CONCEPTO
    ,SUM(CASE WHEN FIFNMONTH = 1 THEN    COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 ENE
    ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 FEB
    ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 MAR
    ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 ABR
    ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 MAY
    ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 JUN
    ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 JUL
    ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 AGO
    ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 SEP
    ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 Oct
    ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 NOV
    ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 DIC
    ,SUM(COALESCE(decimal(FIFNVALMOV,20),0))*1000 TOTAL

    FROM 

    PRODFINA.FNCPXRESCP 

                        WHERE FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
                        AND FIFNYEAR = " + anio + @"
                        AND FIFNSTSEG = 1
                        AND FIFNSTATUS = 1

                        union ALL
                        Select
                        1 IDNOMBREGRUPO
                        ,'AG' NOMBREGRUPO
                        ,22 IDCONCEPTOORDEN
                        ,689 IdConcepto
                        ,'UTILIDAD FINAL ANTES DE IMPUESTOS ACUMULADA' CONCEPTO
                        ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                        ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                        ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                        ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                        ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                        ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                        ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                        ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                        ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                        ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                        ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                        ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                        ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                        FROM PRODFINA.FNDRESMEN
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD = 1055


    UNION ALL


   SELECT 

    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,23 IDCONCEPTOORDEN
    ,690 IdConcepto
    ,'CAPITAL INVERTIDO' CONCEPTO,
    SUM(CASE WHEN FICOMES =  1  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS ENE,
    SUM(CASE WHEN FICOMES =  2  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS FEB,
    SUM(CASE WHEN FICOMES =  3  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS MAR,
    SUM(CASE WHEN FICOMES =  4  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS ABR,
    SUM(CASE WHEN FICOMES =  5  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS MAY,
    SUM(CASE WHEN FICOMES =  6  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS JUN,
    SUM(CASE WHEN FICOMES =  7  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS JUL,
    SUM(CASE WHEN FICOMES =  8  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS AGO,
    SUM(CASE WHEN FICOMES =  9  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS SEP,
    SUM(CASE WHEN FICOMES =  10  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS Oct,
    SUM(CASE WHEN FICOMES =  11  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS NOV,
    SUM(CASE WHEN FICOMES =  12  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS DIC,
    SUM(COALESCE(ABS(decimal(SALDO,20)),0)) AS TOTAL

    FROM (
   SELECT  PERIO.FICOMES , SUM(FDCOSALDOF) SALDO FROM PRODCONT.COESLDCT SALDOS
   INNER JOIN PRODCONT.COCPERIO PERIO ON PERIO.FDCOIDPERI = SALDOS.FDCOIDPERI
   WHERE FICOIDCTA IN(SELECT FIFNIDCTA FROM PRODFINA.FNDCFGCPT WHERE FIFNCPTD IN(81,82,83) AND FIFNIDCIAU = FICOIDCIAU) AND FICOIDCIAU IN(" + Id_Agencia + ')' + @" AND PERIO.FICOANIO = " + anio + @"
  GROUP BY PERIO.FICOMES
  )

    UNION ALL

    Select
    2 IDNOMBREGRUPO
    ,'RENDIMIENTO' NOMBREGRUPO
    ,24 IDCONCEPTOORDEN
    ,691 IdConcepto
    ,'RENDIMIENTO ANTES DE FIDEICOMISO MES' CONCEPTO
    ,SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) ENE
    ,SUM(CASE WHEN FIFNMONTH = 2 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) FEB
    ,SUM(CASE WHEN FIFNMONTH = 3 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) MAR
    ,SUM(CASE WHEN FIFNMONTH = 4 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) ABR
    ,SUM(CASE WHEN FIFNMONTH = 5 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) MAY
    ,SUM(CASE WHEN FIFNMONTH = 6 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) JUN
    ,SUM(CASE WHEN FIFNMONTH = 7 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) JUL
    ,SUM(CASE WHEN FIFNMONTH = 8 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) AGO
    ,SUM(CASE WHEN FIFNMONTH = 9 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) SEP
    ,SUM(CASE WHEN FIFNMONTH = 10 THEN  decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) Oct
    ,SUM(CASE WHEN FIFNMONTH = 11 THEN  decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) NOV
    ,SUM(CASE WHEN FIFNMONTH = 12 THEN  decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) DIC
    ,SUM(decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0)) TOTAL

    FROM PRODFINA.FNDRESMEN   left join (SELECT 

    1 IDNOMBREGRUPO
    ,'AG' NOMBREGRUPO
    ,23 IDCONCEPTOORDEN
    ,1 IdConcepto
    ,'CAPITAL INVERTIDO' CONCEPTO,
    FICOMES,
    SUM(CASE WHEN FICOMES =  1  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS ENE,
    SUM(CASE WHEN FICOMES =  2  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS FEB,
    SUM(CASE WHEN FICOMES =  3  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS MAR,
    SUM(CASE WHEN FICOMES =  4  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS ABR,
    SUM(CASE WHEN FICOMES =  5  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS MAY,
    SUM(CASE WHEN FICOMES =  6  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS JUN,
    SUM(CASE WHEN FICOMES =  7  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS JUL,
    SUM(CASE WHEN FICOMES =  8  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS AGO,
    SUM(CASE WHEN FICOMES =  9  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS SEP,
    SUM(CASE WHEN FICOMES =  10  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS Oct,
    SUM(CASE WHEN FICOMES =  11  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS NOV,
    SUM(CASE WHEN FICOMES =  12  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS DIC,
    SUM(COALESCE(ABS(decimal(SALDO,20)),0)) AS TOTAL

    FROM (
   SELECT  PERIO.FICOMES , SUM(FDCOSALDOF) SALDO FROM PRODCONT.COESLDCT SALDOS
   INNER JOIN PRODCONT.COCPERIO PERIO ON PERIO.FDCOIDPERI = SALDOS.FDCOIDPERI
   WHERE FICOIDCTA IN(SELECT FIFNIDCTA FROM PRODFINA.FNDCFGCPT WHERE FIFNCPTD IN(81,82,83) AND FIFNIDCIAU = FICOIDCIAU) AND FICOIDCIAU IN(" + Id_Agencia + ')' + @" AND PERIO.FICOANIO = " + anio + @"
  GROUP BY PERIO.FICOMES
  ) GROUP BY FICOMES) rnd  on FIFNMONTH = rnd.FICOMES
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD = 1051



    union ALL
    Select
    2 IDNOMBREGRUPO
    ,'RENDIMIENTO' NOMBREGRUPO
    ,25 IDCONCEPTOORDEN
    ,692 IdConcepto
    ,'RENDIMIENTO ANTES DE FIDEICOMISO ACUMULADO' CONCEPTO
    ,SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) ENE
    ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) FEB
    ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) MAR
    ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) ABR
    ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) MAY
    ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) JUN
    ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) JUL
    ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) AGO
    ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) SEP
    ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) Oct
    ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) NOV
    ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) DIC
    ,SUM(decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0))) TOTAL

    FROM PRODFINA.FNDRESMEN   left join (SELECT  
                                        FICOMES
                                        ,1 IDNOMBREGRUPO
                                        ,'AG' NOMBREGRUPO
                                        ,23 IDCONCEPTOORDEN
                                        ,'CAPITAL INVERTIDO' CONCEPTO,
                                        SUM((COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0))) AS Total
                                        FROM (
                                        SELECT  

                                                            PERI.FDCOIDPERI,
                                                            PERI.FICOMES,
                                                            PERI.FICOANIO,
                                                            CONFIGREP.FIFNIDCIAU,
                                                            CONFIGREP.FIFNIDRPT,
                                                            REPORTES.FSFNRPTNM,
                                                            CONFIGREP.FIFNIDAGRP,
                                                            AGRPGENERAL.FSFNIDAGRP,
                                                            CONFIGREP.FIFNIDGRPM,
                                                            AGRPMAESTRO.FSFNAGRPM,CONFIGREP.FIFNIDGRP AS GRUPO,
                                                            NOMBREGRUPO.FSFNGRPNM AS NOMBREGRUPO,
                                                            CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                                                            CONCEPTOS.FSFNNCPT AS CONCEPTO,
                                                            COALESCE(CONFIG.FIFNIDCART,0) IDCART,
                                                            COALESCE(CONFIG.FIFNIDCTA,0) IDCUENTA
                                                            FROM 
                                                            PRODCONT.COCPERIO PERI,
                                                            PRODFINA.FNDCFREPCP CONFIGREP 
                                                            LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP 
                                                            LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM  
                                                            LEFT JOIN PRODFINA.FNCGRPCT NOMBREGRUPO ON CONFIGREP.FIFNIDGRP = NOMBREGRUPO.FIFNIDGRP  
                                                            LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1 
                                                            LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT 
                                                            LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT = CONFIG.FIFNCPTD  
                                                            WHERE 
                                                            PERI.FICOANIO = " + anio + @"
                                                            AND CONFIGREP.FIFNIDRPT = 4 
                                                            AND CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                                            AND CONFIG.FIFNSTATUS = 1
                                                            AND CONFIGREP.FIFNIDCPT BETWEEN 81 AND 83
                                                            ) TB1 
                                                            LEFT JOIN PRODCXC.CCESALCA CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU 
                                                            AND CARTERAS.FICCIDCART = TB1.IDCART 
                                                            AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                                                            AND CARTERAS.FICCSTATUS = 1
                                                            LEFT JOIN PRODCONT.COESLDCT CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI  
                                                            AND CTAS.FICOIDTCCT = 1 
                                                            AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU 
                                                            AND CTAS.FICOIDCTA = TB1.IDCUENTA 
                                                            AND CTAS.FICOSTATUS = 1 
                                                            group by  FICOMES) rnd  on FIFNMONTH = rnd.FICOMES
                        WHERE 
                        FIFNYEAR = " + anio + @"
                        AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                        AND FIFNCPTD = 1051



    UNION ALL

    Select
    3 IDNOMBREGRUPO
    ,'PRODUCTIVIDAD' NOMBREGRUPO
    ,26 IDCONCEPTOORDEN
    ,693 IdConcepto
    ,'PRODUCTIVIDAD DEL MES' CONCEPTO
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 1    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 1 AND T1.FIFNCPTD = 1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) ENE
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 2    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 2 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) FEB
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 3    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 3 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) MAR
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 4    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 4 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) ABR
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 5    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 5 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) MAY
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 6    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 6 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) JUN
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 7    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 7 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) JUL
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 8    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 8 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) AGO
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 9    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 9 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) SEP
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 10   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 10 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) OCT
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 11   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 11 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) NOV
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 12   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 12 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) DIC
    ,DECIMAL((SUM(CASE WHEN T1.FIFNCPTD  = 1051 THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN  T1.FIFNCPTD =  1000  THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) TOTAL

 
                        FROM PRODFINA.FNDRESMEN T1 LEFT JOIN  
                        PRODFINA.FNDRPTMST T2 ON T1.FIFNYEAR = T2.FIFNYEAR
                        AND T1.FIFNMONTH = T2.FIFNMONTH AND T1.FIFNIDCIAU = T2.FIFNIDCIAU 
 
                        WHERE  
                        T1.FIFNYEAR = " + anio + @" 
                        AND T1.FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
                        AND T1.FIFNCPTD IN  (1000,1051) 
                        AND T1.FIFNMONTH IS NOT NULL

    UNION ALL

    Select
    3 IDNOMBREGRUPO
    ,'PRODUCTIVIDAD' NOMBREGRUPO
    ,27 IDCONCEPTOORDEN
    ,694 IdConcepto
    ,'PRODUCTIVIDAD DEL EJERCICIO' CONCEPTO
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 1     AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 1 AND T1.FIFNCPTD = 1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) ENE
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 2    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 2 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) FEB
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 3    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 3 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) MAR
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 4    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 4 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) ABR
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 5    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 5 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) MAY
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 6    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 6 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) JUN
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 7    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 7 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) JUL
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 8    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 8 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) AGO
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 9    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 9 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) SEP
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 10   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 10 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) OCT
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 11   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 11 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) NOV
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 12   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 12 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) DIC
    ,DECIMAL((SUM(CASE WHEN T1.FIFNCPTD  <= 1051 THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN  T1.FIFNCPTD =  1000  THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) TOTAL

 
                        FROM PRODFINA.FNDRESMEN T1 LEFT JOIN  
                        PRODFINA.FNDRPTMST T2 ON T1.FIFNYEAR = T2.FIFNYEAR
                        AND T1.FIFNMONTH = T2.FIFNMONTH AND T1.FIFNIDCIAU = T2.FIFNIDCIAU 
 
                        WHERE  
                        T1.FIFNYEAR = " + anio + @" 
                        AND T1.FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
                        AND T1.FIFNCPTD IN  (1000,1051) 
) vp order by vp.IdConcepto";

            return dbCnx.GetDataTable(query);
        }

        public DataTable GetAGWebV2(string Id_Agencia)
        {
            string query = @"select 
              1 INDX
             -- @AG IDAGENCIA
              ,vp.IDNOMBREGRUPO
              ,vp.NOMBREGRUPO
              ,vp.IdConcepto
              ,vp.CONCEPTO
              ,ifnull(vp.ENE,0) ENE
              ,ifnull(vp.FEB,0) FEB
              ,ifnull(vp.MAR,0) MAR
              ,ifnull(vp.ABR,0) ABR
              ,ifnull(vp.MAY,0) MAY
              ,ifnull(vp.JUN,0) JUN
              ,ifnull(vp.JUL,0) JUL
              ,ifnull(vp.AGO,0) AGO
              ,ifnull(vp.SEP,0) SEP
              ,ifnull(vp.Oct,0) OCT
              ,ifnull(vp.NOV,0) NOV
              ,ifnull(vp.DIC,0) DIC
              ,ifnull(vp.Total,0) TOTAL
          
          
              from (
          
          
               Select
           
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,1 IDCONCEPTOORDEN
              , 668 IdConcepto 
              ,'VENTAS AGENCIA MES' CONCEPTO
              ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
              ,SUM(CASE WHEN FIFNMONTH = 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
              ,SUM(CASE WHEN FIFNMONTH = 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
              ,SUM(CASE WHEN FIFNMONTH = 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
              ,SUM(CASE WHEN FIFNMONTH = 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
              ,SUM(CASE WHEN FIFNMONTH = 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
              ,SUM(CASE WHEN FIFNMONTH = 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
              ,SUM(CASE WHEN FIFNMONTH = 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
              ,SUM(CASE WHEN FIFNMONTH = 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
              ,SUM(CASE WHEN FIFNMONTH = 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
              ,SUM(CASE WHEN FIFNMONTH = 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
              ,SUM(CASE WHEN FIFNMONTH = 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
              ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
          
              FROM PRODFINA.FNDRESMEN
              WHERE 
              FIFNYEAR = " + anio + @"
              AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
              AND FIFNCPTD = 1000
              AND FIFNSTATUS = 1
              
          
              UNION ALL
          
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,2 IDCONCEPTOORDEN
              ,669 IdConcepto
              ,'VENTAS AGENCIA ACUMULADO' CONCEPTO
              ,SUM(CASE WHEN FIFNMONTH = 1 THEN    COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
              ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
              ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
              ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
              ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
              ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
              ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
              ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
              ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
              ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
              ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
              ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
              ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
              FROM PRODFINA.FNDRESMEN
              WHERE 
              FIFNYEAR = " + anio + @"
              AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
              AND FIFNCPTD = 1000
              AND FIFNSTATUS = 1
          
              union ALL
          
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,3 IDCONCEPTOORDEN
              ,670 IdConcepto
              ,'VENTAS FIDEICOMISO MES' CONCEPTO,
              SUM(CASE WHEN FICOMES = 1 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) ENE,
              SUM(CASE WHEN FICOMES = 2 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) FEB,
              SUM(CASE WHEN FICOMES = 3 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) MAR,
              SUM(CASE WHEN FICOMES = 4 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) ABR,
              SUM(CASE WHEN FICOMES = 5 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) MAY,
              SUM(CASE WHEN FICOMES = 6 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) JUN,
              SUM(CASE WHEN FICOMES = 7 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) JUL,
              SUM(CASE WHEN FICOMES = 8 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) AGO,
              SUM(CASE WHEN FICOMES = 9 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) SEP,
              SUM(CASE WHEN FICOMES = 10 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) OCT,
              SUM(CASE WHEN FICOMES = 11 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) NOV,
              SUM(CASE WHEN FICOMES = 12 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) DIC
              ,SUM(CASE WHEN FICOMES <= 12 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) TOTAL
          
              FROM (SELECT PERI.FICOMES,
                     
                     (sum(case when CTADESC.FSCOCUENTA IN ('400000010005') then DECIMAL(((CTAS.FDCOTOTCAR) - (CTAS.FDCOTOTABO)),20,0) else 0 end)                                                                                                                                                                                    
                     ) INGRESOS_AUTOS_NUEVOS
              
                     FROM  PRODCONT.COESLDCT CTAS
                     left join  PRODCONT.COCATCTS CTADESC ON CTAS.FICOIDCTA = CTADESC.FICOIDCTA 
                     LEFT JOIN  PRODCONT.COCPERIO PERI on CTAS.FDCOIDPERI = PERI.FDCOIDPERI  
                     LEFT JOIN (SELECT  CIA.FIGEIDCIAU,CIA.FSGERAZSOC,CIA.FSGESIGCIA,CIA.FSGECALLE,CIA.FIGEIDCLAS,CIA.FIGEIDDIVI,CIA.FIGEIDMARC,MARCA.FNMARCCLA,MARCA.FSMARCDES,MARCA.FNMARCASO,MARCA.FIANSTATU,MARCA.FIANIDCLAS
                                FROM  PRODGRAL.GECCIAUN CIA 
                                LEFT JOIN  PRODAUT.ANCMARCA MARCA  
                                ON CIA.FIGEIDCLAS = MARCA.FIANIDCLAS 
                                and CIA.FIGEIDMARC = MARCA.FNMARCCLA) MARCAS ON MARCAS.FIGEIDCIAU = CTAS.FICOIDCIAU 
                      WHERE  
                      CTAS.FICOIDTCCT = 1 
                      AND CTAS.FICOIDCIAU IN(" + Id_Agencia + ')' + @"
                      AND CTAS.FICOSTATUS = 1
                      AND FICOANIO = " + anio + @"
                      GROUP BY
                      PERI.FICOMES
                      ORDER BY FICOMES ASC)
              UNION ALL
          
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,4 IDCONCEPTOORDEN
              ,671 IdConcepto
              ,'VENTAS FIDEICOMISO ACUMULADO' CONCEPTO,
              SUM(CASE WHEN FICOMES = 1 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) ENE,
              SUM(CASE WHEN FICOMES <= 2 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) FEB,
              SUM(CASE WHEN FICOMES <= 3 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) MAR,
              SUM(CASE WHEN FICOMES <= 4 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) ABR,
              SUM(CASE WHEN FICOMES <= 5 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) MAY,
              SUM(CASE WHEN FICOMES <= 6 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) JUN,
              SUM(CASE WHEN FICOMES <= 7 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) JUL,
              SUM(CASE WHEN FICOMES <= 8 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) AGO,
              SUM(CASE WHEN FICOMES <= 9 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) SEP,
              SUM(CASE WHEN FICOMES <= 10 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) OCT,
              SUM(CASE WHEN FICOMES <= 11 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) NOV,
              SUM(CASE WHEN FICOMES <= 12 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) DIC
              ,SUM(CASE WHEN FICOMES <= 12 THEN ABS(DECIMAL(COALESCE(INGRESOS_AUTOS_NUEVOS,0),20))  ELSE 0 END) TOTAL
          
              FROM (SELECT PERI.FICOMES,
                     
                     (sum(case when CTADESC.FSCOCUENTA IN ('400000010005') then DECIMAL(((CTAS.FDCOTOTCAR) - (CTAS.FDCOTOTABO)),20,0) else 0 end)                                                                                                                                                                                    
                     ) INGRESOS_AUTOS_NUEVOS
              
                     FROM  PRODCONT.COESLDCT CTAS
                     left join  PRODCONT.COCATCTS CTADESC ON CTAS.FICOIDCTA = CTADESC.FICOIDCTA 
                     LEFT JOIN  PRODCONT.COCPERIO PERI on CTAS.FDCOIDPERI = PERI.FDCOIDPERI  
                     LEFT JOIN (SELECT  CIA.FIGEIDCIAU,CIA.FSGERAZSOC,CIA.FSGESIGCIA,CIA.FSGECALLE,CIA.FIGEIDCLAS,CIA.FIGEIDDIVI,CIA.FIGEIDMARC,MARCA.FNMARCCLA,MARCA.FSMARCDES,MARCA.FNMARCASO,MARCA.FIANSTATU,MARCA.FIANIDCLAS
                                FROM  PRODGRAL.GECCIAUN CIA 
                                LEFT JOIN  PRODAUT.ANCMARCA MARCA  
                                ON CIA.FIGEIDCLAS = MARCA.FIANIDCLAS 
                                and CIA.FIGEIDMARC = MARCA.FNMARCCLA) MARCAS ON MARCAS.FIGEIDCIAU = CTAS.FICOIDCIAU 
                      WHERE  
                      CTAS.FICOIDTCCT = 1 
                      AND CTAS.FICOIDCIAU IN(" + Id_Agencia + ')' + @"
                      AND CTAS.FICOSTATUS = 1
                      AND FICOANIO = " + anio + @"
                      GROUP BY
                      PERI.FICOMES
                      ORDER BY FICOMES ASC)
              UNION ALL 
          
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,5 IDCONCEPTOORDEN
              ,672 IdConcepto
              ,'TOTAL VENTAS CON FIDEICOMISO ACUMULADO' CONCEPTO
              ,SUM(CASE WHEN T1.FIFNMONTH  =  1  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH =  1  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) ENE
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 2  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 2  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) FEB
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 3  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 3  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) MAR
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 4  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 4  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) ABR
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 5  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 5  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) MAY
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 6  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 6  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) JUN
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 7  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 7  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) JUL
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 8  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 8  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) AGO
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 9  AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 9  THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) SEP
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 10 AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 10 THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) OCT
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 11 AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 11 THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) NOV
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 12 AND T1.FIFNCPTD = 1000 THEN COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM((CASE WHEN T2.FIFNMONTH <= 12 THEN COALESCE(decimal(T2.FIFNVTAAUF,20,0),0) ELSE 0 END )) DIC
              ,SUM(CASE WHEN T1.FIFNCPTD = 1000 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END) + SUM(COALESCE(decimal(T2.FIFNVTAAUF,20,0),0)) TOTAL
          
              FROM PRODFINA.FNDRESMEN T1 LEFT JOIN 
              PRODFINA.FNDRPTMST T2 ON T1.FIFNYEAR = T2.FIFNYEAR 
              AND T1.FIFNMONTH = T2.FIFNMONTH AND T1.FIFNIDCIAU = T2.FIFNIDCIAU
          
              WHERE 
              T1.FIFNYEAR = " + anio + @"
              AND T1.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
              AND FIFNCPTD = 1000
              AND T1.FIFNSTATUS = 1
              union ALL
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,6 IDCONCEPTOORDEN
              ,673 IdConcepto
              ,'UNIDADES NUEVAS VENDIDAS AGENCIA MES' CONCEPTO
              ,SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 ENE
              ,SUM(CASE WHEN FIFNMONTH = 2 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 FEB
              ,SUM(CASE WHEN FIFNMONTH = 3 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 MAR
              ,SUM(CASE WHEN FIFNMONTH = 4 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 ABR
              ,SUM(CASE WHEN FIFNMONTH = 5 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 MAY
              ,SUM(CASE WHEN FIFNMONTH = 6 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 JUN
              ,SUM(CASE WHEN FIFNMONTH = 7 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 JUL
              ,SUM(CASE WHEN FIFNMONTH = 8 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 AGO
              ,SUM(CASE WHEN FIFNMONTH = 9 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 SEP
              ,SUM(CASE WHEN FIFNMONTH = 10 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 Oct
              ,SUM(CASE WHEN FIFNMONTH = 11 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 NOV
              ,SUM(CASE WHEN FIFNMONTH = 12 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END )*1000 DIC
              ,SUM(decimal(FIFNVALUE,20,0))*1000 TOTAL
          
              FROM PRODFINA.FNDRESMEN
              WHERE 
              FIFNYEAR = " + anio + @"
              AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
              AND FIFNCPTD = 1001
              AND FIFNSTATUS = 1
              union ALL
          
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,7 IDCONCEPTOORDEN
              ,674 IdConcepto
              ,'UNIDADES NUEVAS VENDIDAS AUTOFIN MES' CONCEPTO
              ,SUM(CASE WHEN MES = 1 THEN   COALESCE(UND,0) ELSE 0 END ) ENE
              ,SUM(CASE WHEN MES = 2 THEN  COALESCE(UND,0) ELSE 0 END ) FEB
              ,SUM(CASE WHEN MES = 3 THEN  COALESCE(UND,0) ELSE 0 END ) MAR
              ,SUM(CASE WHEN MES = 4 THEN  COALESCE(UND,0) ELSE 0 END ) ABR
              ,SUM(CASE WHEN MES = 5 THEN  COALESCE(UND,0) ELSE 0 END ) MAY
              ,SUM(CASE WHEN MES = 6 THEN  COALESCE(UND,0) ELSE 0 END ) JUN
              ,SUM(CASE WHEN MES = 7 THEN  COALESCE(UND,0) ELSE 0 END ) JUL
              ,SUM(CASE WHEN MES = 8 THEN  COALESCE(UND,0) ELSE 0 END ) AGO
              ,SUM(CASE WHEN MES = 9 THEN  COALESCE(UND,0) ELSE 0 END ) SEP
              ,SUM(CASE WHEN MES = 10 THEN COALESCE(UND,0)  ELSE 0 END ) Oct
              ,SUM(CASE WHEN MES = 11 THEN COALESCE(UND,0)  ELSE 0 END ) NOV
              ,SUM(CASE WHEN MES = 12 THEN COALESCE(UND,0)  ELSE 0 END ) DIC
              ,SUM(COALESCE(UND,0)) TOTAL
              FROM (
              SELECT * FROM(
                  SELECT  COALESCE(FICAIDCIAU,FIBIIDCIAU) IDCIAU, COALESCE(AÑO,YEAR(FFBIFECHA)) AÑO, COALESCE(MES,MONTH(FFBIFECHA)) MES, COALESCE(UNIDADES,0) UND
                  FROM   
                  (
                      SELECT  FICAIDCIAU, AÑO, MES, SUM(UNIDADES) UNIDADES , sum(FDCASUBTOT) VENTA, SUM(FDCATOTAL) VENTA_F
                      FROM   (
                          SELECT  FA.FICAIDCIAU,YEAR(FFCAFECHA) AÑO, MONTH(FFCAFECHA) MES, sum(FA.FDCASUBTOT) FDCASUBTOT ,sum(FA.FDCATOTAL) FDCATOTAL ,  COUNT(*) UNIDADES
                          FROM PRODCAJA.CAEFACTU FA
                          INNER JOIN        PRODCAJA.CAEFACAN FN
                          ON         FA.FICAIDCIAU=FN.FICAIDCIAU AND FA.FICAIDFACT=FN.FICAIDFACT
                          LEFT JOIN           PRODCAJA.CAENOTCR NC
                          ON         FA.FICAIDCIAU=NC.FICAIDCIAU AND FA.FICAIDFACT=NC.FICAIDFACT
                          AND       NC.FICAIDTINC=3 AND FICAIDESTA=4
                          WHERE FA.FICAIDTIFA=3 
                                          AND (NC.FICAIDFACT IS NOT NULL OR FA.FICASTATUS=1) 
                                          AND YEAR(FFCAFECHA) IN (" + anio + @")                                                                --AND MONTH(FFCAFECHA) IN (8)
                                          AND FSCAMARCA NOT IN ('AIMA','BAJAJ') 
                                          AND FA.FICAIDCIAU IN(" + Id_Agencia + ')' + @"
                                          AND FICAIDTIVT IN (3)                                                       
          
                          GROUP BY  FA.FICAIDCIAU, YEAR(FFCAFECHA) ,MONTH(FFCAFECHA)
                          UNION
                          SELECT  FA.FICAIDCIAU, YEAR(FFCAFECHAE) AÑO, MONTH(FFCAFECHAE) MES,sum(FA.FDCASUBTOT)*-1 FDCASUBTOT, sum(FA.FDCATOTAL)*-1 FDCATOTAL, COUNT(*)*-1 UNIDADES
                          FROM   PRODCAJA.CAEFACTU FA
                          INNER JOIN        PRODCAJA.CAEFACAN FN
                          ON         FA.FICAIDCIAU=FN.FICAIDCIAU AND FA.FICAIDFACT=FN.FICAIDFACT
                          INNER JOIN        PRODCAJA.CAENOTCR NC
                          ON         FA.FICAIDCIAU=NC.FICAIDCIAU AND FA.FICAIDFACT=NC.FICAIDFACT
                          AND       NC.FICAIDTINC=3 AND FICAIDESTA=4
                          WHERE NC.FICASTATUS=1
                                          AND FA.FICAIDTIFA=3 
                                          AND YEAR(FFCAFECHAE) IN (" + anio + @")
                                          AND FSCAMARCA NOT IN ('AIMA','BAJAJ') 
                                          AND FA.FICAIDCIAU IN(" + Id_Agencia + ')' + @"
                                          AND FICAIDTIVT IN (3)
                          GROUP BY          FA.FICAIDCIAU, YEAR(FFCAFECHAE) ,MONTH(FFCAFECHAE)
                      ) X
                      GROUP BY          FICAIDCIAU, AÑO, MES
                  )S
                  FULL JOIN (
                              SELECT * FROM PRODBUIN.BIDINDAG WHERE FIBIIDINDI=52 AND FIBIIDTIPO=2 AND FIBIIDAREA=1 AND YEAR(FFBIFECHA) IN (" + anio + @")
                          ) IND
                  ON  FICAIDCIAU=FIBIIDCIAU AND AÑO=YEAR(FFBIFECHA) AND MES=MONTH(FFBIFECHA)
                  WHERE 
                  FICAIDCIAU IN(" + Id_Agencia + ')' + @"
                  ) WHERE AÑO = " + anio + @"
              )
          
              UNION ALL 
          
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,8 IDCONCEPTOORDEN
              ,675 IdConcepto
              ,'UNIDADES NUEVAS VENDIDAS TAXIS BAM  MES' CONCEPTO
              ,SUM(0) ENE
              ,SUM(0) FEB
              ,SUM(0) MAR
              ,SUM(0) ABR
              ,SUM(0) MAY
              ,SUM(0) JUN
              ,SUM(0) JUL
              ,SUM(0) AGO
              ,SUM(0) SEP
              ,SUM(0)  Oct
              ,SUM(0)  NOV
              ,SUM(0) DIC
              ,SUM(0) TOTAL
              FROM PRODFINA.FNDRPTMST
              WHERE 
              FIFNYEAR = " + anio + @"
              AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
          
          
              UNION ALL
           SELECT 1 IDNOMBREGRUPO
                  ,'AG' NOMBREGRUPO
                  ,9 IDCONCEPTOORDEN
                  ,676 IdConcepto
                  ,'TOTAL UNIDADES NUEVAS VENDIDAS MES' CONCEPTO
                  ,SUM(ENE)*1000 ENE
                  ,SUM(FEB)*1000 FEB
                  ,SUM(MAR)*1000 MAR
                  ,SUM(ABR)*1000 ABR
                  ,SUM(MAY)*1000 MAY
                  ,SUM(JUN)*1000 JUN
                  ,SUM(JUL)*1000 JUL
                  ,SUM(AGO)*1000 AGO
                  ,SUM(SEP)*1000 SEP
                  ,SUM(OCT)*1000 OCT
                  ,SUM(NOV)*1000 NOV
                  ,SUM(DIC)*1000 DIC
                  ,SUM(TOTAL)*1000 TOTAL
                  FROM (
                      SELECT 
           SUM(ENE) ENE
          , SUM(FEB) FEB
          , SUM(MAR) MAR
          , SUM(ABR) ABR
          , SUM(MAY) MAY
          , SUM(JUN) JUN
          , SUM(JUL) JUL
          , SUM(AGO) AGO
          , SUM(SEP) SEP
          , SUM(OCT) OCT
          , SUM(NOV) NOV
          , SUM(DIC) DIC
          , SUM(TOTAL) TOTAL
          FROM(
          SELECT 
                              SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ENE
                              ,SUM(CASE WHEN FIFNMONTH = 2 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) FEB
                              ,SUM(CASE WHEN FIFNMONTH = 3 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAR
                              ,SUM(CASE WHEN FIFNMONTH = 4 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ABR
                              ,SUM(CASE WHEN FIFNMONTH = 5 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAY
                              ,SUM(CASE WHEN FIFNMONTH = 6 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUN
                              ,SUM(CASE WHEN FIFNMONTH = 7 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUL
                              ,SUM(CASE WHEN FIFNMONTH = 8 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) AGO
                              ,SUM(CASE WHEN FIFNMONTH = 9 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) SEP
                              ,SUM(CASE WHEN FIFNMONTH = 10 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) OCT
                              ,SUM(CASE WHEN FIFNMONTH = 11 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) NOV
                              ,SUM(CASE WHEN FIFNMONTH = 12 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) DIC
                              ,SUM(decimal(FIFNVALUE,20,0)) TOTAL
                             FROM PRODFINA.FNDVIHIST WHERE FIFNIDCIAU IN(" + Id_Agencia + ')' + @" AND FIFNYEAR = " + anio + @"  AND FIFNCAMPO IN (1,4) AND FIFNSTATUS = 1 GROUP BY FIFNMONTH
                   ) )
          
              UNION ALL
          
          
          
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,10 IDCONCEPTOORDEN
              ,677 IdConcepto
              ,'TOTAL UNIDADES NUEVAS VENDIDAS ACUMULADO' CONCEPTO
              ,SUM(ENE)*1000 ENE
                  ,SUM(FEB)*1000 FEB
                  ,SUM(MAR)*1000 MAR
                  ,SUM(ABR)*1000 ABR
                  ,SUM(MAY)*1000 MAY
                  ,SUM(JUN)*1000 JUN
                  ,SUM(JUL)*1000 JUL
                  ,SUM(AGO)*1000 AGO
                  ,SUM(SEP)*1000 SEP
                  ,SUM(OCT)*1000 OCT
                  ,SUM(NOV)*1000 NOV
                  ,SUM(DIC)*1000 DIC
                  ,SUM(TOTAL)*1000 TOTAL
                  FROM (
                      SELECT 
           SUM(ENE) ENE
          , SUM(FEB) FEB
          , SUM(MAR) MAR
          , SUM(ABR) ABR
          , SUM(MAY) MAY
          , SUM(JUN) JUN
          , SUM(JUL) JUL
          , SUM(AGO) AGO
          , SUM(SEP) SEP
          , SUM(OCT) OCT
          , SUM(NOV) NOV
          , SUM(DIC) DIC
          , SUM(TOTAL) TOTAL
          FROM(
          SELECT 
                              SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ENE
                              ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) FEB
                              ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAR
                              ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ABR
                              ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAY
                              ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUN
                              ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUL
                              ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) AGO
                              ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) SEP
                              ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) OCT
                              ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) NOV
                              ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) DIC
                              ,SUM(decimal(FIFNVALUE,20,0)) TOTAL
                             FROM PRODFINA.FNDVIHIST WHERE FIFNIDCIAU IN(" + Id_Agencia + ')' + @" AND FIFNYEAR = " + anio + @"  AND FIFNCAMPO IN (1,4) AND FIFNSTATUS = 1 GROUP BY FIFNMONTH
                   ) )
          
              UNION ALL
          
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,11 IDCONCEPTOORDEN
              ,678 IdConcepto
              ,'PROMEDIO UTILIDAD BRUTA POR UNIDAD NUEVA VENDIDA ACUMULADO' CONCEPTO
              ,SUM(CASE WHEN T1.FIFNMONTH  =  1  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH  = 1  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) ENE
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 2  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 2  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) FEB
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 3  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 3  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) MAR
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 4  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 4  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) ABR
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 5  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 5  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) MAY
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 6  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 6  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) JUN
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 7  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 7  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) JUL
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 8  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 8  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END)  AGO
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 9  AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 9  AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END)  SEP
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 10 AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 10 AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) Oct
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 11 AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 11 AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) NOV
              ,SUM(CASE WHEN T1.FIFNMONTH  <= 12 AND T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNMONTH <= 12 AND T1.FIFNCPTD = 1001 THEN  COALESCE(decimal    (T1.FIFNVALUE,20,0),0)        ELSE 0 END) DIC
              ,SUM(CASE WHEN T1.FIFNCPTD = 1004 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0)  ELSE 0 END) / SUM(CASE WHEN T1.FIFNCPTD =  1001 THEN  COALESCE(decimal(T1.FIFNVALUE,20,0),0) ELSE 0 END ) TOTAL
          
           
                                  FROM PRODFINA.FNDRESMEN T1 LEFT JOIN  
                                  PRODFINA.FNDRPTMST T2 ON T1.FIFNYEAR = T2.FIFNYEAR
                                  AND T1.FIFNMONTH = T2.FIFNMONTH AND T1.FIFNIDCIAU = T2.FIFNIDCIAU 
           
                                  WHERE  
                                  T1.FIFNYEAR = " + anio + @" 
                                  AND T1.FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
                                  AND T1.FIFNCPTD IN  (1001,1004) 
           AND T1.FIFNSTATUS = 1
              union ALL
          
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,12 IDCONCEPTOORDEN
              ,679 IdConcepto
              ,'UTILIDAD BRUTA PROMEDIO ACUMULADO' CONCEPTO
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH = 1    AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH = 1    AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  ENE  
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 2   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 2   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  FEB  
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 3   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 3   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  MAR
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 4   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 4   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  ABR
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 5   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 5   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  MAY
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 6   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 6   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  JUN
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 7   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 7   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  JUL
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 8   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 8   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  AGO
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 9   AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 9   AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  SEP
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 10  AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 10  AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  Oct
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 11  AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 11  AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  NOV
              ,decimal(SUM(CASE WHEN T1.FIFNMONTH <= 12  AND T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNMONTH <= 12  AND T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20)   ELSE   0       END ),20)  DIC
              ,decimal(SUM(CASE WHEN T1.FIFNCPTD =  1004 THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END),20) / decimal(SUM(CASE WHEN T1.FIFNCPTD =   1002 THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END ),20)  TOTAL
          
          
          
                                  FROM PRODFINA.FNDRESMEN T1
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD IN  (1004,1002)
           AND T1.FIFNSTATUS = 1
          
              union ALL
              Select
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,13 IDCONCEPTOORDEN
              ,680 IdConcepto
              ,'GASTO CORPORATIVO DEL MES' CONCEPTO
              ,SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ENE
              ,SUM(CASE WHEN FIFNMONTH = 2 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) FEB
              ,SUM(CASE WHEN FIFNMONTH = 3 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAR
              ,SUM(CASE WHEN FIFNMONTH = 4 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) ABR
              ,SUM(CASE WHEN FIFNMONTH = 5 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) MAY
              ,SUM(CASE WHEN FIFNMONTH = 6 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUN
              ,SUM(CASE WHEN FIFNMONTH = 7 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) JUL
              ,SUM(CASE WHEN FIFNMONTH = 8 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) AGO
              ,SUM(CASE WHEN FIFNMONTH = 9 THEN   decimal(FIFNVALUE,20,0) ELSE 0 END ) SEP
              ,SUM(CASE WHEN FIFNMONTH = 10 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) Oct
              ,SUM(CASE WHEN FIFNMONTH = 11 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) NOV
              ,SUM(CASE WHEN FIFNMONTH = 12 THEN  decimal(FIFNVALUE,20,0) ELSE 0 END ) DIC
              ,SUM(decimal(FIFNVALUE,20,0)) TOTAL
          
                                  FROM PRODFINA.FNDRESMEN
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD = 1050
           AND FIFNSTATUS = 1
                                  union ALL
                                  Select
                                  1 IDNOMBREGRUPO
                                  ,'AG' NOMBREGRUPO
                                  ,14 IDCONCEPTOORDEN
                                  ,681 IdConcepto
                                  ,'UTILIDAD ANTES DEL FIDEICOMISO MES' CONCEPTO
                                  ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                                  ,SUM(CASE WHEN FIFNMONTH = 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                                  ,SUM(CASE WHEN FIFNMONTH = 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                                  ,SUM(CASE WHEN FIFNMONTH = 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                                  ,SUM(CASE WHEN FIFNMONTH = 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                                  ,SUM(CASE WHEN FIFNMONTH = 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                                  ,SUM(CASE WHEN FIFNMONTH = 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                                  ,SUM(CASE WHEN FIFNMONTH = 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                                  ,SUM(CASE WHEN FIFNMONTH = 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                                  ,SUM(CASE WHEN FIFNMONTH = 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                                  ,SUM(CASE WHEN FIFNMONTH = 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                                  ,SUM(CASE WHEN FIFNMONTH = 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                                  ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                                  FROM PRODFINA.FNDRESMEN
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD = 1051
           AND FIFNSTATUS = 1
              union ALL
              
              SELECT 
              TB1.IDNOMBREGRUPO
              ,TB1.NOMBREGRUPO
              ,TB1.IDCONCEPTOORDEN
              ,TB1.IdConcepto
              ,TB1.CONCEPTO
              ,IFNULL(TB1.ENE,0) ENE
              ,IFNULL(TB1.FEB,0) FEB
              ,IFNULL(TB1.MAR,0) MAR
              ,IFNULL(TB1.ABR,0) ABR
              ,IFNULL(TB1.MAY,0) MAY
              ,IFNULL(TB1.JUN,0) JUN
              ,IFNULL(TB1.JUL,0) JUL
              ,IFNULL(TB1.AGO,0) AGO
              ,IFNULL(TB1.SEP,0) SEP
              ,IFNULL(TB1.OCT,0) OCT
              ,IFNULL(TB1.NOV,0) NOV
              ,IFNULL(TB1.DIC,0) DIC
              ,IFNULL(TB1.TOTAL,0) TOTAL
          
          
          
          
              FROM (
          
          
                                  Select
                                                              1 IDNOMBREGRUPO
                                                              ,'AG' NOMBREGRUPO
                                                              ,15 IDCONCEPTOORDEN
                                                              ,682 IdConcepto
                                                              ,'UTILIDAD FIDEICOMISO MES' CONCEPTO
                                                              ,SUM(CASE WHEN FIFNMONTH = 1  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                                                              ,SUM(CASE WHEN FIFNMONTH = 2  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                                                              ,SUM(CASE WHEN FIFNMONTH = 3  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                                                              ,SUM(CASE WHEN FIFNMONTH = 4  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                                                              ,SUM(CASE WHEN FIFNMONTH = 5  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                                                              ,SUM(CASE WHEN FIFNMONTH = 6  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                                                              ,SUM(CASE WHEN FIFNMONTH = 7  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                                                              ,SUM(CASE WHEN FIFNMONTH = 8  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                                                              ,SUM(CASE WHEN FIFNMONTH = 9  THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                                                              ,SUM(CASE WHEN FIFNMONTH = 10  THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                                                              ,SUM(CASE WHEN FIFNMONTH = 11  THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                                                              ,SUM(CASE WHEN FIFNMONTH = 12  THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                                                              ,SUM(COALESCE(IFNULL(FIFNVALUE,0),0)) TOTAL
                                                              FROM PRODFINA.FNDRESMEN
                                                              WHERE 
                                                              FIFNYEAR = " + anio + @"
                                                              AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                                              AND FIFNCPTD = 1052
          
              ) TB1
          
              UNION ALL 
          
              SELECT 
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,16 IDCONCEPTOORDEN
              ,683 IdConcepto
              ,'PARTIDAS EXTRAORDINARIAS MES' CONCEPTO
              ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 ENE
              ,SUM(CASE WHEN FIFNMONTH = 2 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 FEB
              ,SUM(CASE WHEN FIFNMONTH = 3 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 MAR
              ,SUM(CASE WHEN FIFNMONTH = 4 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 ABR
              ,SUM(CASE WHEN FIFNMONTH = 5 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 MAY
              ,SUM(CASE WHEN FIFNMONTH = 6 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 JUN
              ,SUM(CASE WHEN FIFNMONTH = 7 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 JUL
              ,SUM(CASE WHEN FIFNMONTH = 8 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 AGO
              ,SUM(CASE WHEN FIFNMONTH = 9 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 SEP
              ,SUM(CASE WHEN FIFNMONTH = 10 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 Oct
              ,SUM(CASE WHEN FIFNMONTH = 11 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 NOV
              ,SUM(CASE WHEN FIFNMONTH = 12 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 DIC
              ,SUM(COALESCE(decimal(FIFNVALMOV,20),0))*1000 TOTAL
          
              FROM 
          
              PRODFINA.FNCPXRESCP 
          
              WHERE FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
              AND FIFNYEAR = " + anio + @"
              AND FIFNSTSEG = 1
              AND FIFNSTATUS = 1
          
                                  union ALL
                                  Select
                                  1 IDNOMBREGRUPO
                                  ,'AG' NOMBREGRUPO
                                  ,17 IDCONCEPTOORDEN
                                  ,684 IdConcepto
                                  ,'UTILIDAD FINAL ANTES DE IMPUESTOS MES' CONCEPTO
                                  ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                                  ,SUM(CASE WHEN FIFNMONTH = 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                                  ,SUM(CASE WHEN FIFNMONTH = 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                                  ,SUM(CASE WHEN FIFNMONTH = 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                                  ,SUM(CASE WHEN FIFNMONTH = 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                                  ,SUM(CASE WHEN FIFNMONTH = 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                                  ,SUM(CASE WHEN FIFNMONTH = 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                                  ,SUM(CASE WHEN FIFNMONTH = 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                                  ,SUM(CASE WHEN FIFNMONTH = 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                                  ,SUM(CASE WHEN FIFNMONTH = 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                                  ,SUM(CASE WHEN FIFNMONTH = 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                                  ,SUM(CASE WHEN FIFNMONTH = 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                                  ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                                  FROM PRODFINA.FNDRESMEN
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD = 1055
          
          
                                  union ALL
                                  Select
                                  1 IDNOMBREGRUPO
                                  ,'AG' NOMBREGRUPO
                                  ,18 IDCONCEPTOORDEN
                                  ,685 IdConcepto
                                  ,'GASTO CORPORATIVO ACUMULADO' CONCEPTO
                                  ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                                  ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                                  ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                                  ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                                  ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                                  ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                                  ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                                  ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                                  ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                                  ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                                  ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                                  ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                                  ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                                  FROM PRODFINA.FNDRESMEN
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD = 1050
          
          
              union ALL
          
              SELECT 
              TB1.IDNOMBREGRUPO
              ,TB1.NOMBREGRUPO
              ,TB1.IDCONCEPTOORDEN
              ,TB1.IdConcepto
              ,TB1.CONCEPTO
              ,IFNULL(TB1.ENE,0) ENE
              ,IFNULL(TB1.FEB,0) FEB
              ,IFNULL(TB1.MAR,0) MAR
              ,IFNULL(TB1.ABR,0) ABR
              ,IFNULL(TB1.MAY,0) MAY
              ,IFNULL(TB1.JUN,0) JUN
              ,IFNULL(TB1.JUL,0) JUL
              ,IFNULL(TB1.AGO,0) AGO
              ,IFNULL(TB1.SEP,0) SEP
              ,IFNULL(TB1.OCT,0) OCT
              ,IFNULL(TB1.NOV,0) NOV
              ,IFNULL(TB1.DIC,0) DIC
              ,IFNULL(TB1.TOTAL,0) TOTAL
          
          
          
          
              FROM (
          
          
          
                                  Select
                                  1 IDNOMBREGRUPO
                                  ,'AG' NOMBREGRUPO
                                  ,19 IDCONCEPTOORDEN
                                  ,686 IdConcepto
                                  ,'UTILIDAD ANTES DE FIDEICOMISO ACUMULADO' CONCEPTO
                                  ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                                  ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                                  ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                                  ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                                  ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                                  ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                                  ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                                  ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                                  ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                                  ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                                  ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                                  ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                                  ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                                  FROM PRODFINA.FNDRESMEN
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD = 1051
          
              ) TB1
          
              union ALL
          
              SELECT 
              TB1.IDNOMBREGRUPO
              ,TB1.NOMBREGRUPO
              ,TB1.IDCONCEPTOORDEN
              ,TB1.IdConcepto
              ,TB1.CONCEPTO
              ,IFNULL(TB1.ENE,0) ENE
              ,IFNULL(TB1.FEB,0) FEB
              ,IFNULL(TB1.MAR,0) MAR
              ,IFNULL(TB1.ABR,0) ABR
              ,IFNULL(TB1.MAY,0) MAY
              ,IFNULL(TB1.JUN,0) JUN
              ,IFNULL(TB1.JUL,0) JUL
              ,IFNULL(TB1.AGO,0) AGO
              ,IFNULL(TB1.SEP,0) SEP
              ,IFNULL(TB1.OCT,0) OCT
              ,IFNULL(TB1.NOV,0) NOV
              ,IFNULL(TB1.DIC,0) DIC
              ,IFNULL(TB1.TOTAL,0) TOTAL
          
                                  FROM (
                                  Select
                                  1 IDNOMBREGRUPO
                                  ,'AG' NOMBREGRUPO
                                  ,20 IDCONCEPTOORDEN
                                  ,687 IdConcepto
                                  ,'UTILIDAD FIDEICOMISO ACUMULADA' CONCEPTO
                                  ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                                  ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                                  ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                                  ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                                  ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                                  ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                                  ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                                  ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                                  ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                                  ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                                  ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                                  ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                                  ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                                  FROM PRODFINA.FNDRESMEN
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD = 1052
                                  ) TB1
          
              UNION ALL 
          
              SELECT 
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,21 IDCONCEPTOORDEN
              ,688 IdConcepto
              ,'PARTIDAS EXTRAORDINARIAS ACUMULADO' CONCEPTO
              ,SUM(CASE WHEN FIFNMONTH = 1 THEN    COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 ENE
              ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 FEB
              ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 MAR
              ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 ABR
              ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 MAY
              ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 JUN
              ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 JUL
              ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 AGO
              ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 SEP
              ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 Oct
              ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 NOV
              ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(decimal(FIFNVALMOV,20),0) ELSE 0 END )*1000 DIC
              ,SUM(COALESCE(decimal(FIFNVALMOV,20),0))*1000 TOTAL
          
              FROM 
          
              PRODFINA.FNCPXRESCP 
          
                                  WHERE FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
                                  AND FIFNYEAR = " + anio + @"
                                  AND FIFNSTSEG = 1
                                  AND FIFNSTATUS = 1
          
                                  union ALL
                                  Select
                                  1 IDNOMBREGRUPO
                                  ,'AG' NOMBREGRUPO
                                  ,22 IDCONCEPTOORDEN
                                  ,689 IdConcepto
                                  ,'UTILIDAD FINAL ANTES DE IMPUESTOS ACUMULADA' CONCEPTO
                                  ,SUM(CASE WHEN FIFNMONTH = 1 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ENE
                                  ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) FEB
                                  ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAR
                                  ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) ABR
                                  ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) MAY
                                  ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUN
                                  ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) JUL
                                  ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) AGO
                                  ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   COALESCE(FIFNVALUE,0) ELSE 0 END ) SEP
                                  ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) Oct
                                  ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) NOV
                                  ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  COALESCE(FIFNVALUE,0) ELSE 0 END ) DIC
                                  ,SUM(COALESCE(FIFNVALUE,0)) TOTAL
                                  FROM PRODFINA.FNDRESMEN
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD = 1055
          
          
              UNION ALL
          
          
             SELECT 
          
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,23 IDCONCEPTOORDEN
              ,690 IdConcepto
              ,'CAPITAL INVERTIDO' CONCEPTO,
              SUM(CASE WHEN FICOMES =  1  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS ENE,
              SUM(CASE WHEN FICOMES =  2  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS FEB,
              SUM(CASE WHEN FICOMES =  3  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS MAR,
              SUM(CASE WHEN FICOMES =  4  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS ABR,
              SUM(CASE WHEN FICOMES =  5  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS MAY,
              SUM(CASE WHEN FICOMES =  6  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS JUN,
              SUM(CASE WHEN FICOMES =  7  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS JUL,
              SUM(CASE WHEN FICOMES =  8  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS AGO,
              SUM(CASE WHEN FICOMES =  9  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS SEP,
              SUM(CASE WHEN FICOMES =  10  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS Oct,
              SUM(CASE WHEN FICOMES =  11  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS NOV,
              SUM(CASE WHEN FICOMES =  12  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS DIC,
              SUM(COALESCE(ABS(decimal(SALDO,20)),0)) AS TOTAL
          
              FROM (
             SELECT  PERIO.FICOMES , SUM(FDCOSALDOF) SALDO FROM PRODCONT.COESLDCT SALDOS
             INNER JOIN PRODCONT.COCPERIO PERIO ON PERIO.FDCOIDPERI = SALDOS.FDCOIDPERI
             WHERE FICOIDCTA IN(SELECT FIFNIDCTA FROM PRODFINA.FNDCFGCPT WHERE FIFNCPTD IN(81,82,83) AND FIFNIDCIAU = FICOIDCIAU) AND FICOIDCIAU IN(" + Id_Agencia + ')' + @" AND PERIO.FICOANIO = " + anio + @"
            GROUP BY PERIO.FICOMES
            )
          
              UNION ALL
          
              Select
              2 IDNOMBREGRUPO
              ,'RENDIMIENTO' NOMBREGRUPO
              ,24 IDCONCEPTOORDEN
              ,691 IdConcepto
              ,'RENDIMIENTO ANTES DE FIDEICOMISO MES' CONCEPTO
              ,SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) ENE
              ,SUM(CASE WHEN FIFNMONTH = 2 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) FEB
              ,SUM(CASE WHEN FIFNMONTH = 3 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) MAR
              ,SUM(CASE WHEN FIFNMONTH = 4 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) ABR
              ,SUM(CASE WHEN FIFNMONTH = 5 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) MAY
              ,SUM(CASE WHEN FIFNMONTH = 6 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) JUN
              ,SUM(CASE WHEN FIFNMONTH = 7 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) JUL
              ,SUM(CASE WHEN FIFNMONTH = 8 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) AGO
              ,SUM(CASE WHEN FIFNMONTH = 9 THEN   decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) SEP
              ,SUM(CASE WHEN FIFNMONTH = 10 THEN  decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) Oct
              ,SUM(CASE WHEN FIFNMONTH = 11 THEN  decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) NOV
              ,SUM(CASE WHEN FIFNMONTH = 12 THEN  decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0) ELSE 0 END ) DIC
              ,SUM(decimal(FIFNVALUE,20,0)/decimal(rnd.Total,20,0)) TOTAL
          
              FROM PRODFINA.FNDRESMEN   left join (SELECT 
          
              1 IDNOMBREGRUPO
              ,'AG' NOMBREGRUPO
              ,23 IDCONCEPTOORDEN
              ,1 IdConcepto
              ,'CAPITAL INVERTIDO' CONCEPTO,
              FICOMES,
              SUM(CASE WHEN FICOMES =  1  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS ENE,
              SUM(CASE WHEN FICOMES =  2  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS FEB,
              SUM(CASE WHEN FICOMES =  3  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS MAR,
              SUM(CASE WHEN FICOMES =  4  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS ABR,
              SUM(CASE WHEN FICOMES =  5  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS MAY,
              SUM(CASE WHEN FICOMES =  6  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS JUN,
              SUM(CASE WHEN FICOMES =  7  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS JUL,
              SUM(CASE WHEN FICOMES =  8  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS AGO,
              SUM(CASE WHEN FICOMES =  9  THEN    COALESCE(ABS(decimal(SALDO,20)),0) END ) AS SEP,
              SUM(CASE WHEN FICOMES =  10  THEN   COALESCE(ABS(decimal(SALDO,20)),0) END ) AS Oct,
              SUM(CASE WHEN FICOMES =  11  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS NOV,
              SUM(CASE WHEN FICOMES =  12  THEN  COALESCE(ABS(decimal(SALDO,20)),0) END ) AS DIC,
              SUM(COALESCE(ABS(decimal(SALDO,20)),0)) AS TOTAL
          
              FROM (
             SELECT  PERIO.FICOMES , SUM(FDCOSALDOF) SALDO FROM PRODCONT.COESLDCT SALDOS
             INNER JOIN PRODCONT.COCPERIO PERIO ON PERIO.FDCOIDPERI = SALDOS.FDCOIDPERI
             WHERE FICOIDCTA IN(SELECT FIFNIDCTA FROM PRODFINA.FNDCFGCPT WHERE FIFNCPTD IN(81,82,83) AND FIFNIDCIAU = FICOIDCIAU) AND FICOIDCIAU IN(" + Id_Agencia + ')' + @" AND PERIO.FICOANIO = " + anio + @"
            GROUP BY PERIO.FICOMES
            ) GROUP BY FICOMES) rnd  on FIFNMONTH = rnd.FICOMES
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD = 1051
          
          
          
              union ALL
              Select
              2 IDNOMBREGRUPO
              ,'RENDIMIENTO' NOMBREGRUPO
              ,25 IDCONCEPTOORDEN
              ,692 IdConcepto
              ,'RENDIMIENTO ANTES DE FIDEICOMISO ACUMULADO' CONCEPTO
              ,SUM(CASE WHEN FIFNMONTH = 1 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) ENE
              ,SUM(CASE WHEN FIFNMONTH <= 2 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) FEB
              ,SUM(CASE WHEN FIFNMONTH <= 3 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) MAR
              ,SUM(CASE WHEN FIFNMONTH <= 4 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) ABR
              ,SUM(CASE WHEN FIFNMONTH <= 5 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) MAY
              ,SUM(CASE WHEN FIFNMONTH <= 6 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) JUN
              ,SUM(CASE WHEN FIFNMONTH <= 7 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) JUL
              ,SUM(CASE WHEN FIFNMONTH <= 8 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) AGO
              ,SUM(CASE WHEN FIFNMONTH <= 9 THEN   decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) SEP
              ,SUM(CASE WHEN FIFNMONTH <= 10 THEN  decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) Oct
              ,SUM(CASE WHEN FIFNMONTH <= 11 THEN  decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) NOV
              ,SUM(CASE WHEN FIFNMONTH <= 12 THEN  decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0)) ELSE 0 END ) DIC
              ,SUM(decimal(FIFNVALUE,20,0)/ABS(decimal(rnd.Total,20,0))) TOTAL
          
              FROM PRODFINA.FNDRESMEN   left join (SELECT  
                                                  FICOMES
                                                  ,1 IDNOMBREGRUPO
                                                  ,'AG' NOMBREGRUPO
                                                  ,23 IDCONCEPTOORDEN
                                                  ,'CAPITAL INVERTIDO' CONCEPTO,
                                                  SUM((COALESCE(CARTERAS.FDCCSALDOF,0) + COALESCE(CTAS.FDCOSALDOF,0))) AS Total
                                                  FROM (
                                                  SELECT  
          
                                                                      PERI.FDCOIDPERI,
                                                                      PERI.FICOMES,
                                                                      PERI.FICOANIO,
                                                                      CONFIGREP.FIFNIDCIAU,
                                                                      CONFIGREP.FIFNIDRPT,
                                                                      REPORTES.FSFNRPTNM,
                                                                      CONFIGREP.FIFNIDAGRP,
                                                                      AGRPGENERAL.FSFNIDAGRP,
                                                                      CONFIGREP.FIFNIDGRPM,
                                                                      AGRPMAESTRO.FSFNAGRPM,CONFIGREP.FIFNIDGRP AS GRUPO,
                                                                      NOMBREGRUPO.FSFNGRPNM AS NOMBREGRUPO,
                                                                      CONFIGREP.FIFNIDCPT AS IDCONCEPTO,
                                                                      CONCEPTOS.FSFNNCPT AS CONCEPTO,
                                                                      COALESCE(CONFIG.FIFNIDCART,0) IDCART,
                                                                      COALESCE(CONFIG.FIFNIDCTA,0) IDCUENTA
                                                                      FROM 
                                                                      PRODCONT.COCPERIO PERI,
                                                                      PRODFINA.FNDCFREPCP CONFIGREP 
                                                                      LEFT JOIN PRODFINA.FNCAGRGRAL AGRPGENERAL ON CONFIGREP.FIFNIDAGRP = AGRPGENERAL.FIFNIDAGRP 
                                                                      LEFT JOIN PRODFINA.FNCAGRUPA AGRPMAESTRO ON CONFIGREP.FIFNIDGRPM = AGRPMAESTRO.FIFNIDGRPM  
                                                                      LEFT JOIN PRODFINA.FNCGRPCT NOMBREGRUPO ON CONFIGREP.FIFNIDGRP = NOMBREGRUPO.FIFNIDGRP  
                                                                      LEFT JOIN PRODFINA.FNCCONCP CONCEPTOS ON CONFIGREP.FIFNIDCPT = CONCEPTOS.FIFNIDCPT  AND CONCEPTOS.FIFNSTATUS = 1 
                                                                      LEFT JOIN PRODFINA.FNCREPORT REPORTES ON CONFIGREP.FIFNIDRPT = REPORTES.FIFNIDRPT 
                                                                      LEFT JOIN PRODFINA.FNDCFGCPT CONFIG ON CONFIG.FIFNIDCIAU = CONFIGREP.FIFNIDCIAU AND CONFIG.FIFNIDRPT = CONFIGREP.FIFNIDRPT  AND CONFIGREP.FIFNIDCPT =        CONFIG.FIFNCPTD    
                                                                      WHERE 
                                                                      PERI.FICOANIO = " + anio + @"
                                                                      AND CONFIGREP.FIFNIDRPT = 4 
                                                                      AND CONFIGREP.FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                                                      AND CONFIG.FIFNSTATUS = 1
                                                                      AND CONFIGREP.FIFNIDCPT BETWEEN 81 AND 83
                                                                      ) TB1 
                                                                      LEFT JOIN PRODCXC.CCESLCAS CARTERAS  ON CARTERAS.FICCIDCIAU = TB1.FIFNIDCIAU 
                                                                      AND CARTERAS.FICCIDCART = TB1.IDCART 
                                                                      AND CARTERAS.FDCCIDPERI = TB1.FDCOIDPERI
                                                                      AND CARTERAS.FICCSTATUS = 1
                                                                      LEFT JOIN PRODCONT.COESLDEX CTAS ON CTAS.FDCOIDPERI = TB1.FDCOIDPERI  
                                                                      AND CTAS.FICOIDTCCT = 1 
                                                                      AND CTAS.FICOIDCIAU = TB1.FIFNIDCIAU 
                                                                      AND CTAS.FICOIDCTA = TB1.IDCUENTA 
                                                                      AND CTAS.FICOSTATUS = 1 
                                                                      group by  FICOMES) rnd  on FIFNMONTH = rnd.FICOMES
                                  WHERE 
                                  FIFNYEAR = " + anio + @"
                                  AND FIFNIDCIAU IN(" + Id_Agencia + ')' + @"
                                  AND FIFNCPTD = 1051
          
          
          
              UNION ALL
          
              Select
              3 IDNOMBREGRUPO
              ,'PRODUCTIVIDAD' NOMBREGRUPO
              ,26 IDCONCEPTOORDEN
              ,693 IdConcepto
              ,'PRODUCTIVIDAD DEL MES' CONCEPTO
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 1    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 1 AND T1.FIFNCPTD = 1000  THEN T1.FIFNVALUE ELSE   0         END)/1000),20) ENE
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 2    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 2 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE  ELSE   0        END)/1000),20) FEB
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 3    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 3 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE  ELSE   0        END)/1000),20) MAR
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 4    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 4 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE  ELSE   0        END)/1000),20) ABR
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 5    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 5 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE  ELSE   0        END)/1000),20) MAY
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 6    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 6 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE  ELSE   0        END)/1000),20) JUN
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 7    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 7 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE  ELSE   0        END)/1000),20) JUL
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 8    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 8 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE  ELSE   0        END)/1000),20) AGO
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 9    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 9 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE  ELSE   0        END)/1000),20) SEP
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 10   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 10 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE   ELSE   0       END)/1000),20) OCT
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 11   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 11 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE   ELSE   0       END)/1000),20) NOV
              ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 12   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 12 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE   ELSE   0       END)/1000),20) DIC
              ,DECIMAL((SUM(CASE WHEN T1.FIFNCPTD  = 1051 THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN  T1.FIFNCPTD =  1000  THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) TOTAL
          
           
                                  FROM PRODFINA.FNDRESMEN T1 LEFT JOIN  
                                  PRODFINA.FNDRPTMST T2 ON T1.FIFNYEAR = T2.FIFNYEAR
                                  AND T1.FIFNMONTH = T2.FIFNMONTH AND T1.FIFNIDCIAU = T2.FIFNIDCIAU 
           
                                  WHERE  
                                  T1.FIFNYEAR = " + anio + @" 
                                  AND T1.FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
                                  AND T1.FIFNCPTD IN  (1000,1051) 
                                  AND T1.FIFNMONTH IS NOT NULL
          
              UNION ALL
          
              Select
    3 IDNOMBREGRUPO
    ,'PRODUCTIVIDAD' NOMBREGRUPO
    ,27 IDCONCEPTOORDEN
    ,694 IdConcepto
    ,'PRODUCTIVIDAD DEL EJERCICIO' CONCEPTO
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH = 1     AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  = 1 AND T1.FIFNCPTD = 1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) ENE
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 2    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 2 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) FEB
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 3    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 3 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) MAR
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 4    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 4 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) ABR
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 5    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 5 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) MAY
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 6    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 6 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) JUN
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 7    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 7 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) JUL
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 8    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 8 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) AGO
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 9    AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 9 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) SEP
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 10   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 10 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) OCT
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 11   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 11 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) NOV
    ,DECIMAL((SUM(CASE WHEN T1.FIFNMONTH <= 12   AND  T1.FIFNCPTD =  1051  THEN  decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN T1.FIFNMONTH  <= 12 AND T1.FIFNCPTD =  1000  THEN T1.FIFNVALUE ELSE 0 END)/1000),20) DIC
    ,DECIMAL((SUM(CASE WHEN T1.FIFNCPTD  <= 1051 THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) / DECIMAL((SUM(CASE WHEN  T1.FIFNCPTD =  1000  THEN decimal(T1.FIFNVALUE,20,0) ELSE 0 END)/1000),20) TOTAL

 
                        FROM PRODFINA.FNDRESMEN T1 LEFT JOIN  
                        PRODFINA.FNDRPTMST T2 ON T1.FIFNYEAR = T2.FIFNYEAR
                        AND T1.FIFNMONTH = T2.FIFNMONTH AND T1.FIFNIDCIAU = T2.FIFNIDCIAU 
 
                        WHERE  
                        T1.FIFNYEAR = " + anio + @" 
                        AND T1.FIFNIDCIAU IN(" + Id_Agencia + ')' + @" 
                        AND T1.FIFNCPTD IN  (1000,1051) 
        ) vp order by vp.IdConcepto";

            return dbCnx.GetDataTable(query);
        }

        public void GeneraExcelSC(List<int> idsAgencia, bool esUnaAgencia)
        {
            List<Agencia> agencias;

            if (idsAgencia == null)
            {
                List<AgenciasReportes> agenciasReportes = AgenciasReportes.Listar(_db, 1); //Todas las agencias
                List<int> aIdAgencias = new List<int>();

                aIdAgencias.AddRange(agenciasReportes.Select(o => o.IdAgencia));
                agencias = Agencia.ListarPorIds(_db, aIdAgencias);
            }
            else
                agencias = Agencia.ListarPorIds(_db, idsAgencia);

            string idAgenciaString = string.Join(", ", agencias.Select(item => item.Id));

            List<ConceptosContables> conceptosV2 = ConceptosContables.ListarSC_V1yV2(_db); //ToDo cambiar
            conceptosV2 = conceptosV2.Where(e => e.Id < 10602).ToList();


            DataTable dtBGExcelV1 = GetExcelV1(13);
            DataTable dtBGExcelV1Acum = GetExcelV1Acumulado(13);
            DataTable dtBGExcelV2 = GetExcelV2(13);
            DataTable dtBGExcelV2Acum = GetExcelV2Acumulado(13);

            int r = 1;
            int c = 1;
            int vExcel = 0;
            int vWeb = 0;
            int intAux = 0;
            DataRow[] drBGExcelV1 = null;
            DataRow[] drBGExcelV1Acum = null;
            DataRow[] drBGExcelV2 = null;
            DataRow[] drBGExcelV2Acum = null;
            DataRow[] drBGWebV1 = null;
            DataRow[] drBGWebV1Acum = null;
            DataRow[] drBGWebV2 = null;
            DataRow[] drBGWebV2Acum = null;

            List<Reporte> listaR = new List<Reporte>();
            List<ReporteGlobal> listaRG = new List<ReporteGlobal>();

            DataTable dtBGWebV1 = new DataTable();
            DataTable dtBGWebV2 = new DataTable();

            //var getResTable = GetAGWebV1("28");

            using (DVAExcel.ExcelWriter eW = new DVAExcel.ExcelWriter(ruta))
            {
                foreach (Agencia agencia in agencias.OrderBy(o => o.Siglas))
                {
                    if (esUnaAgencia)
                    {
                        dtBGWebV1 = GetSCWebV1(idAgenciaString);
                        dtBGWebV2 = GetSCWebV2(idAgenciaString);
                    }
                    else
                    {
                        dtBGWebV1 = GetSCWebV1(agencia.Id.ToString());
                        dtBGWebV2 = GetSCWebV2(agencia.Id.ToString());
                    }


                    if ((agencia.Id == 590) || (agencia.Id == 583) || (agencia.Id == 301) || (agencia.Id == 583) || (agencia.Id == 563) || (agencia.Id == 593)
                        || (agencia.Id == 592) || (agencia.Id == 594) || (agencia.Id == 100))
                        continue;

                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("[SIGLAS]: " + agencia.Siglas);
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");
                    Console.WriteLine("******************************************************************************************************");


                    int countRow = 0;
                    foreach (ConceptosContables concepto in conceptosV2)
                    {
                        Reporte rep = new Reporte();

                        r++;

                        rep.ID_CONCEPTO = concepto.Id;            // ID Concepto                
                        rep.CONCEPTO = concepto.NombreConcepto;   // Nombre Concepto



                        drBGExcelV1 = dtBGExcelV1.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drBGExcelV1Acum = dtBGExcelV1Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);


                        //    drBGWebV1 = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);
                        //    drBGWebV1Acum = dtBGWebV1.Select("IDCONCEPTO = " + concepto.Id);



                        drBGExcelV2 = dtBGExcelV2.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);
                        drBGExcelV2Acum = dtBGExcelV2Acum.Select("ID_AGENCIA = " + agencia.Id + " AND ID_CONCEPTO = " + concepto.Id);


                        //drBGWebV2 = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);
                        //drBGWebV2Acum = dtBGWebV2.Select("IDCONCEPTO = " + concepto.Id);


                        if (drBGExcelV1.Length != 0)
                        {
                        vExcel = Convert.ToInt32(drBGExcelV1[0]["VALOR"]);
                        rep.EXCEL_V1 = vExcel;
                        vExcel = Convert.ToInt32(drBGExcelV1Acum[0]["VALOR"]);
                        rep.EXCEL_V1_ACUM = vExcel;
                        }


                        vWeb = 0;



                        if (concepto.Id >= 10560 && concepto.Id <= 10566) // Documentos por Cobrar
                        {
                            //vWeb += Convert.ToInt32(dr["FIFNDPCMP"]); 

                            var valoresColumna = dtBGWebV1.AsEnumerable()
                              .Select(row => row.Field<decimal>("FIFNDPCMP"))
                              .ToList();

                            if (valoresColumna.Count != 0 && valoresColumna.Count == 7)
                            {
                                var posicion1 = valoresColumna[countRow];
                                vWeb += Convert.ToInt32(posicion1);
                                countRow++;
                            }
                            if (countRow == 7)
                            {
                                countRow = 0;
                            }



                        }
                        else if (concepto.Id >= 10567 && concepto.Id <= 10573)
                        {
                            //vWeb += Convert.ToInt32(dr["FIFNDPCMPN"]);

                            var valoresColumna_ = dtBGWebV1.AsEnumerable()
                              .Select(row => row.Field<decimal>("FIFNDPCMPN"))
                              .ToList();

                            if (valoresColumna_.Count != 0 && valoresColumna_.Count == 7)
                            {
                                var posicion1 = valoresColumna_[countRow];
                                vWeb += Convert.ToInt32(posicion1);
                                countRow++;
                            }


                            if (countRow == 7)
                            {
                                countRow = 0;
                            }

                        }
                        else if (concepto.Id >= 10574 && concepto.Id <= 10580)
                        {
                            //vWeb += Convert.ToInt32(dr["FIFNDPCMPF"]);

                            var valoresColumna_ = dtBGWebV1.AsEnumerable()
                            .Select(row => row.Field<decimal>("FIFNDPCMPF"))
                            .ToList();

                            if (valoresColumna_.Count != 0 && valoresColumna_.Count == 7)
                            {
                                var posicion1 = valoresColumna_[countRow];
                                vWeb += Convert.ToInt32(posicion1);
                                countRow++;
                            }


                            if (countRow == 7)
                            {
                                countRow = 0;
                            }
                        }
                        else if (concepto.Id >= 10581 && concepto.Id <= 10587)
                        {
                            //vWeb += Convert.ToInt32(dr["FIFNDPCMPS"]);
                            var valoresColumna_ = dtBGWebV1.AsEnumerable()
                            .Select(row => row.Field<decimal>("FIFNDPCMPS"))
                            .ToList();

                            if (valoresColumna_.Count != 0 && valoresColumna_.Count == 7)
                            {
                                var posicion1 = valoresColumna_[countRow];
                                vWeb += Convert.ToInt32(posicion1);
                                countRow++;
                            }


                            if (countRow == 7)
                            {
                                countRow = 0;
                            }
                        }
                        else if (concepto.Id >= 10588 && concepto.Id <= 10594)
                        {
                            //vWeb += Convert.ToInt32(dr["FIFNDPCMPR"]);

                            var valoresColumna_ = dtBGWebV1.AsEnumerable()
                         .Select(row => row.Field<decimal>("FIFNDPCMPR"))
                         .ToList();

                            if (valoresColumna_.Count != 0 && valoresColumna_.Count == 7)
                            {
                                var posicion1 = valoresColumna_[countRow];
                                vWeb += Convert.ToInt32(posicion1);
                                countRow++;
                            }

                            if (countRow == 7)
                            {
                                countRow = 0;
                            }
                        }
                        else if (concepto.Id >= 10595 && concepto.Id <= 10601)
                        {
                            //vWeb += Convert.ToInt32(dr["FIFNDPCMPP"]);
                            var valoresColumna_ = dtBGWebV1.AsEnumerable()
                    .Select(row => row.Field<decimal>("FIFNDPCMPP"))
                    .ToList();

                            if (valoresColumna_.Count != 0 && valoresColumna_.Count == 7)
                            {
                                var posicion1 = valoresColumna_[countRow];
                                vWeb += Convert.ToInt32(posicion1);
                                countRow++;
                            }

                            
                            if (countRow == 7)
                            {
                                countRow = 0;
                            }
                        }

                        rep.WEB_V1 = vWeb;




                        vWeb = 0;
                        //rep.WEB_V1_ACUM = vWeb * cambioSigo;
                        rep.WEB_V1_ACUM = vWeb;




                        rep.DIFF_V1 = rep.EXCEL_V1 - rep.WEB_V1;
                        rep.DIFF_V1_ACUM = rep.EXCEL_V1_ACUM - rep.WEB_V1_ACUM;

                        if (drBGExcelV2.Length != 0)
                        {
                            vExcel = Convert.ToInt32(drBGExcelV2[0]["VALOR"]);
                            rep.EXCEL_V2 = vExcel;
                            vExcel = Convert.ToInt32(drBGExcelV2Acum[0]["VALOR"]);
                            rep.EXCEL_V2_ACUM = vExcel;
                            vWeb = 0;
                        }

                        vWeb = 0;
                        //rep.WEB_V2_ACUM = vWeb * cambioSigo;

                        rep.WEB_V2_ACUM = vWeb;




                        rep.DIFF_V2 = rep.EXCEL_V2 - rep.WEB_V2;
                        rep.DIFF_V2_ACUM = rep.EXCEL_V2_ACUM - rep.WEB_V2_ACUM;

                        Console.WriteLine("[CONCEPTO]: " + concepto.Id + "_" + concepto.NombreConcepto +
                            " [EXCEL_V1]: " + rep.EXCEL_V1 +
                            " [WEB_V1]: " + rep.WEB_V1 +
                            " [DIFF_V1]: " + rep.DIFF_V1 +
                            " [EXCEL_V1_ACUMULADO]: " + rep.EXCEL_V1_ACUM +
                            " [WEB_V1_ACUMULADO]: " + rep.WEB_V1_ACUM +
                            " [DIFF_V1_ACUMULADO]: " + rep.DIFF_V1_ACUM +
                            " [EXCEL_V2]: " + rep.EXCEL_V2 +
                            " [WEB_V2]: " + rep.WEB_V2 +
                            " [DIFF_V2]: " + rep.DIFF_V2 +
                            " [EXCEL_V2_ACUMULADO]: " + rep.EXCEL_V2_ACUM +
                            " [WEB_V2_ACUMULADO]: " + rep.WEB_V2_ACUM +
                            " [DIFF_V2_ACUMULADO]: " + rep.DIFF_V2_ACUM);

                        listaR.Add(rep);
                    }

                    var dtToExcel = listaR.ToDataTable();
                    eW.WriteDataTable(listaR.ToDataTable(), agencia.Siglas);
                    
                    listaR = new List<Reporte>();
                }

                eW.Dispose();
            }
        }
        public DataTable GetSCWebV1(string Id_Agencia)
        {
            var anios = "";
            var mesAnterior_ = 0;

            var anioAnte = "";
            var mesActual_ = mes;
            var anioAnterior_ = 0;
            var anioActual_ = anio;

            if (mes == 1)
            {
                var anioAnt = anio - 1;
                anioAnte = anioAnt.ToString();
                anios = anio.ToString() + "," + anioAnte;
                mesAnterior_ = 12;
                anioAnterior_ = anioAnt;
            }
            else
            {
                anioAnterior_ = anio;
                mesAnterior_ = mes - 1;
                anios = anio.ToString();
            }

            #region query
            string query = @"SELECT 
                           " + Id_Agencia + @" ID_AGENCIA,
                           " + anio + @" FIFNYEAR,
                           " + mes + @"  FIFNMONTH,
                           CASE 1
                           WHEN 1 THEN 'ENE'
                           WHEN 2 THEN 'FEB'
                           WHEN 3 THEN 'MAR'
                           WHEN 4 THEN 'ABR'
                           WHEN 5 THEN 'MAY'
                           WHEN 6 THEN 'JUN'
                           WHEN 7 THEN 'JUL'
                           WHEN 8 THEN 'AGO'
                           WHEN 9 THEN 'SEP'
                           WHEN 10 THEN 'OCT'
                           WHEN 11 THEN 'NOV'
                           ELSE 'DIC' END AS FSFNMONTH,
                           T1.FIFNCPTD ID_CONCEPTO,                  
                           T1.FSFNNCPT FSFNNCPT, 
                           SUM(TP.DocsMesAnterior)/1000 FIFNDPCMA, 
                           SUM(TP.DocsMesActual)/1000        FIFNDPCMP, 
                           SUM(TP.AutosNuevMesAnterior)/1000 FIFNDPCMAN,
                           SUM(TP.AutosNuevMesActual)/1000   FIFNDPCMPN, 
                           SUM(TP.AutoFinMesAnterior)/1000   FIFNDPCMAF, 
                           SUM(TP.AutoFinMesActual)/1000     FIFNDPCMPF,
                           SUM(TP.ServMesAnterior)/1000    FIFNDPCMAS,
                           SUM(TP.ServMesActual)/1000        FIFNDPCMPS,
                           SUM(TP.RefacMesAnterior)/1000  FIFNDPCMAR,
                           SUM(TP.RefacMesActual)/1000       FIFNDPCMPR,
                           SUM(TP.PlantaMesAnterior)/1000 FIFNDPCMAP,
                           SUM(TP.PlantaMesActual)/1000      FIFNDPCMPP,
                           SUM(TP.JuridicoPenalMesAnterior)/1000 FIFNDJPMAN,
                           SUM(TP.JuridicoPenalMesActual)/1000      FIFNDJPMAC
                           FROM (VALUES (1,'CARTERA ACTIVA')
                                  , (2,'MONTO POR VENCER')
                                  , (3,'VENCIDO DE 1 A 30 DIAS')
                                  , (4,'VENCIDO DE 31 A 60 DIAS')
                                  , (5,'VENCIDO DE 61 A 90 DIAS')
                                  , (6,'VENCIDO DE 91 A 120 DIAS')
                                  , (7,'VENCIDO MAS DE 120 DIAS')
                                  , (8,'CARTERA PASIVA')
                            ) T1 (FIFNCPTD, FSFNNCPT) 
                            LEFT JOIN (                       
                            SELECT *
                           FROM (VALUES
                                    (1,1), (1,2), (1,3), (1,4), (1,5), (1,6), (1,7), (1,8), (1,9), (1,10), (1,11), (1,12)
                                   ,(2,1), (2,2), (2,3), (2,4), (2,5), (2,11) 
                                   ,(3,6)
                                   ,(4,7)
                                   ,(5,8)
                                   ,(6,9)
                                   ,(7,10), (7,12)
                            ) T2 (FIFNCPTDA, FSFNNCPT)) T2 ON T1.FIFNCPTD = T2.FIFNCPTDA --SE SUSTITUIRA POR UNA TABLA
                         INNER JOIN (  
                                          
                         SELECT 
                         FICCIDANCO,
                         COALESCE(DocsMesAnterior,0) DocsMesAnterior, 
                         COALESCE(DocsMesActual,0) DocsMesActual, 
                         COALESCE(AutosNuevMesAnterior,0) AutosNuevMesAnterior, 
                         COALESCE(AutosNuevMesActual,0) AutosNuevMesActual, 
                         COALESCE(AutoFinMesAnterior,0) AutoFinMesAnterior, 
                         COALESCE(AutoFinMesActual,0) AutoFinMesActual, 
                         COALESCE(ServMesAnterior,0) ServMesAnterior, 
                         COALESCE(ServMesActual,0) ServMesActual,
                         COALESCE(RefacMesAnterior,0) RefacMesAnterior,
                         COALESCE(RefacMesActual,0) RefacMesActual,
                         COALESCE(PlantaMesAnterior,0) PlantaMesAnterior,
                         COALESCE(PlantaMesActual,0) PlantaMesActual,
                         COALESCE(JuridicoPenalMesAnterior,0) JuridicoPenalMesAnterior,
                         COALESCE(JuridicoPenalMesActual,0) JuridicoPenalMesActual                                
                         
                         FROM (
                         
                         SELECT ANTIG.FICCIDANCO ,SALCO.FICCIDCIAU, SALCO.FICCIDCART, PERSO.FDPEIDTCTE, /*--Importe inicial*/ 
                         
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @"  AND CARTE.FICCIDCART IN (1,2,3,4,151) THEN SUM(ANTIG.X1) END DocsMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (1,2,3,4,151) THEN SUM(ANTIG.X1) END DocsMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (6,8,9,10,11,12,13,66,69) THEN SUM(ANTIG.X1) END AutosNuevMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (6,8,9,10,11,12,13,66,69) THEN SUM(ANTIG.X1) END AutosNuevMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (7) THEN SUM(ANTIG.X1) END AutoFinMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (7) THEN SUM(ANTIG.X1) END AutoFinMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (14,15,109) THEN SUM(ANTIG.X1) END ServMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (14,15,109) THEN SUM(ANTIG.X1) END ServMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (17,18,19) THEN SUM(ANTIG.X1) END RefacMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (17,18,19) THEN SUM(ANTIG.X1) END RefacMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (16,23,24,26,27,28,57,58,59,60) THEN SUM(ANTIG.X1) END PlantaMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (16,23,24,26,27,28,57,58,59,60) THEN SUM(ANTIG.X1) END PlantaMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (143,144,5) THEN SUM(ANTIG.X1) END JuridicoPenalMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (143,144,5) THEN SUM(ANTIG.X1) END JuridicoPenalMesActual 
                                            
                         
                         FROM PRODCXC.CCESALCO SALCO 
                         INNER JOIN PRODPERS.CTEPERSO PERSO 
                         ON SALCO.FICCIDPERS = PERSO.FDPEIDPERS 
                         INNER JOIN PRODGRAL.GECCIAUN CIAUN 
                         ON SALCO.FICCIDCIAU = CIAUN.FIGEIDCIAU 
                         INNER JOIN PRODCXC.CCCARTER CARTE 
                         ON SALCO.FICCIDCART = CARTE.FICCIDCART 
                         LEFT JOIN PRODPERS.CTDPERMO PERMO 
                         ON SALCO.FICCIDPERS = PERMO.FDPMIDPERS 
                         LEFT JOIN PRODPERS.CTCPERFI PERFI 
                         ON SALCO.FICCIDPERS = PERFI.FDPFIDPERS                     
                         
                         LEFT JOIN ( 
                         SELECT FDPDIDPERS, MIN(FDPDCODPOS) FDPDCODPOS
                         FROM PRODPERS.CTDRPDOM 
                         WHERE FBCTDEFAUL = 1 AND
                         FDPDESTATU = 1 GROUP BY FDPDIDPERS) RPDOM 
                         ON SALCO.FICCIDPERS = RPDOM.FDPDIDPERS 
                         
                         INNER JOIN PRODCONT.COCPERIO CPERIO 
                         ON SALCO.FFCCFECSAC BETWEEN CPERIO.FFCOFECINI AND CPERIO.FFCOFECFIN 
                         LEFT JOIN PRODCXC.CCESALPE SALPE 
                         ON SALCO.FICCIDCIAU = SALPE.FICCIDCIAU 
                         AND SALCO.FICCIDPERS = SALPE.FDCCIDPERS 
                         AND SALCO.FICCIDCART = SALPE.FICCIDCART 
                         AND CPERIO.FDCOIDPERI = SALPE.FDCCIDPERI 
                         INNER JOIN 
                         (SELECT ANSAC.FICCIDANCO,ANSAC.FICCIDCIAU, ANSAC.FICCIDSACO, ANSAC.FFCCFECSAC, ANSAC.FDCCIMPORP X1 
                         FROM PRODCXC.CCDANSAC ANSAC 
                         GROUP BY ANSAC.FICCIDCIAU, ANSAC.FICCIDSACO, ANSAC.FFCCFECSAC,ANSAC.FICCIDANCO,ANSAC.FDCCIMPORP) ANTIG 
                         ON SALCO.FICCIDCIAU = ANTIG.FICCIDCIAU 
                         AND SALCO.FICCIDSACO = ANTIG.FICCIDSACO 
                         AND SALCO.FFCCFECSAC = ANTIG.FFCCFECSAC                     
                         WHERE 
                         SALCO.FFCCFECSAC = CPERIO.FFCOFECFIN 
                         AND CPERIO.FICOANIO in (" + anios + ')' + @"      
                         AND SALCO.FICCIDCIAU NOT IN (1,91,92,65,245,15, 58,111,210,121,244) 
                         AND SALCO.FICCIDCIAU IN (" + Id_Agencia + ')' + @"
                         GROUP BY ANTIG.FICCIDANCO,SALCO.FICCIDCIAU, SALCO.FICCIDPERS, SALCO.FICCIDCART, PERSO.FDPEIDTCTE, CIAUN.FSGERAZCOM, CIAUN.FIGEIDMARC, CARTE.FSCCDESCAR, CARTE.FICCIDCART,CPERIO.FICOMES,CPERIO.FICOANIO)    TPRINTER                            
                         ) TP ON TP.FICCIDANCO IN (T2.FSFNNCPT)             
                         GROUP BY T1.FIFNCPTD, T1.FSFNNCPT   
                                        
                         ORDER BY T1.FIFNCPTD";
            #endregion


            return dbCnx.GetDataTable(query);
        }
        public DataTable GetSCWebV2(string Id_Agencia)
        {
            //mes = 1;
            var anios = "";
            var mesAnterior_ = 0;

            var anioAnte = "";
            var mesActual_ = mes;
            var anioAnterior_ = 0;
            var anioActual_ = anio;

            if (mes == 1)
            {
                var anioAnt = anio - 1;
                anioAnte = anioAnt.ToString();
                anios = anio.ToString() + "," + anioAnte;
                mesAnterior_ = 12;
                anioAnterior_ = anioAnt;
            }
            else
            {
                anioAnterior_ = anio;
                mesAnterior_ = mes - 1;
                anios = anio.ToString();
            }

            #region query V2
            string query = @"SELECT 
                           " + Id_Agencia + @" ID_AGENCIA,
                           " + anio + @" FIFNYEAR,
                           " + mes + @"  FIFNMONTH,
                           CASE 1
                           WHEN 1 THEN 'ENE'
                           WHEN 2 THEN 'FEB'
                           WHEN 3 THEN 'MAR'
                           WHEN 4 THEN 'ABR'
                           WHEN 5 THEN 'MAY'
                           WHEN 6 THEN 'JUN'
                           WHEN 7 THEN 'JUL'
                           WHEN 8 THEN 'AGO'
                           WHEN 9 THEN 'SEP'
                           WHEN 10 THEN 'OCT'
                           WHEN 11 THEN 'NOV'
                           ELSE 'DIC' END AS FSFNMONTH,
                           T1.FIFNCPTD ID_CONCEPTO,                  
                           T1.FSFNNCPT FSFNNCPT, 
                           SUM(TP.DocsMesAnterior)/1000 FIFNDPCMA, 
                           SUM(TP.DocsMesActual)/1000        FIFNDPCMP, 
                           SUM(TP.AutosNuevMesAnterior)/1000 FIFNDPCMAN,
                           SUM(TP.AutosNuevMesActual)/1000   FIFNDPCMPN, 
                           SUM(TP.AutoFinMesAnterior)/1000   FIFNDPCMAF, 
                           SUM(TP.AutoFinMesActual)/1000     FIFNDPCMPF,
                           SUM(TP.ServMesAnterior)/1000    FIFNDPCMAS,
                           SUM(TP.ServMesActual)/1000        FIFNDPCMPS,
                           SUM(TP.RefacMesAnterior)/1000  FIFNDPCMAR,
                           SUM(TP.RefacMesActual)/1000       FIFNDPCMPR,
                           SUM(TP.PlantaMesAnterior)/1000 FIFNDPCMAP,
                           SUM(TP.PlantaMesActual)/1000      FIFNDPCMPP,
                           SUM(TP.JuridicoPenalMesAnterior)/1000 FIFNDJPMAN,
                           SUM(TP.JuridicoPenalMesActual)/1000      FIFNDJPMAC
                           FROM (VALUES (1,'CARTERA ACTIVA')
                                  , (2,'MONTO POR VENCER')
                                  , (3,'VENCIDO DE 1 A 30 DIAS')
                                  , (4,'VENCIDO DE 31 A 60 DIAS')
                                  , (5,'VENCIDO DE 61 A 90 DIAS')
                                  , (6,'VENCIDO DE 91 A 120 DIAS')
                                  , (7,'VENCIDO MAS DE 120 DIAS')
                                  , (8,'CARTERA PASIVA')
                            ) T1 (FIFNCPTD, FSFNNCPT) 
                            LEFT JOIN (                       
                            SELECT *
                           FROM (VALUES
                                    (1,1), (1,2), (1,3), (1,4), (1,5), (1,6), (1,7), (1,8), (1,9), (1,10), (1,11), (1,12)
                                   ,(2,1), (2,2), (2,3), (2,4), (2,5), (2,11) 
                                   ,(3,6)
                                   ,(4,7)
                                   ,(5,8)
                                   ,(6,9)
                                   ,(7,10), (7,12)
                            ) T2 (FIFNCPTDA, FSFNNCPT)) T2 ON T1.FIFNCPTD = T2.FIFNCPTDA --SE SUSTITUIRA POR UNA TABLA
                         INNER JOIN (  
                                          
                         SELECT 
                         FICCIDANCO,
                         COALESCE(DocsMesAnterior,0) DocsMesAnterior, 
                         COALESCE(DocsMesActual,0) DocsMesActual, 
                         COALESCE(AutosNuevMesAnterior,0) AutosNuevMesAnterior, 
                         COALESCE(AutosNuevMesActual,0) AutosNuevMesActual, 
                         COALESCE(AutoFinMesAnterior,0) AutoFinMesAnterior, 
                         COALESCE(AutoFinMesActual,0) AutoFinMesActual, 
                         COALESCE(ServMesAnterior,0) ServMesAnterior, 
                         COALESCE(ServMesActual,0) ServMesActual,
                         COALESCE(RefacMesAnterior,0) RefacMesAnterior,
                         COALESCE(RefacMesActual,0) RefacMesActual,
                         COALESCE(PlantaMesAnterior,0) PlantaMesAnterior,
                         COALESCE(PlantaMesActual,0) PlantaMesActual,
                         COALESCE(JuridicoPenalMesAnterior,0) JuridicoPenalMesAnterior,
                         COALESCE(JuridicoPenalMesActual,0) JuridicoPenalMesActual                                
                         
                         FROM (
                         
                         SELECT ANTIG.FICCIDANCO ,SALCO.FICCIDCIAU, SALCO.FICCIDCART, PERSO.FDPEIDTCTE, /*--Importe inicial*/ 
                         
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @"  AND CARTE.FICCIDCART IN (1,2,3,4,144,151) THEN SUM(ANTIG.X1) END DocsMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (1,2,3,4,144,151) THEN SUM(ANTIG.X1) END DocsMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (6,8,9,10,11,12,13,66,69) THEN SUM(ANTIG.X1) END AutosNuevMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (6,8,9,10,11,12,13,66,69) THEN SUM(ANTIG.X1) END AutosNuevMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (7) THEN SUM(ANTIG.X1) END AutoFinMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (7) THEN SUM(ANTIG.X1) END AutoFinMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (14,15,109) THEN SUM(ANTIG.X1) END ServMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (14,15,109) THEN SUM(ANTIG.X1) END ServMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (17,18,19) THEN SUM(ANTIG.X1) END RefacMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (17,18,19) THEN SUM(ANTIG.X1) END RefacMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (16,23,24,26,27,28,57,58,59,60) THEN SUM(ANTIG.X1) END PlantaMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (16,23,24,26,27,28,57,58,59,60) THEN SUM(ANTIG.X1) END PlantaMesActual,
                         CASE WHEN CPERIO.FICOMES = " + mesAnterior_ + @" AND CPERIO.FICOANIO = " + anioAnterior_ + @" AND CARTE.FICCIDCART IN (143,5) THEN SUM(ANTIG.X1) END JuridicoPenalMesAnterior,
                         CASE WHEN CPERIO.FICOMES = " + mesActual_ + @" AND CPERIO.FICOANIO = " + anioActual_ + @" AND CARTE.FICCIDCART IN (143,5) THEN SUM(ANTIG.X1) END JuridicoPenalMesActual 
                                            
                         
                         FROM PRODCXC.CCESALCD  SALCO 
                         INNER JOIN PRODPERS.CTEPERSO PERSO 
                         ON SALCO.FICCIDPERS = PERSO.FDPEIDPERS 
                         INNER JOIN PRODGRAL.GECCIAUN CIAUN 
                         ON SALCO.FICCIDCIAU = CIAUN.FIGEIDCIAU 
                         INNER JOIN PRODCXC.CCCARTER CARTE 
                         ON SALCO.FICCIDCART = CARTE.FICCIDCART 
                         LEFT JOIN PRODPERS.CTDPERMO PERMO 
                         ON SALCO.FICCIDPERS = PERMO.FDPMIDPERS 
                         LEFT JOIN PRODPERS.CTCPERFI PERFI 
                         ON SALCO.FICCIDPERS = PERFI.FDPFIDPERS                     
                         
                         LEFT JOIN ( 
                         SELECT FDPDIDPERS, MIN(FDPDCODPOS) FDPDCODPOS
                         FROM PRODPERS.CTDRPDOM 
                         WHERE FBCTDEFAUL = 1 AND
                         FDPDESTATU = 1 GROUP BY FDPDIDPERS) RPDOM 
                         ON SALCO.FICCIDPERS = RPDOM.FDPDIDPERS 
                         
                         INNER JOIN PRODCONT.COCPERIO CPERIO 
                         ON SALCO.FFCCFECSAC BETWEEN CPERIO.FFCOFECINI AND CPERIO.FFCOFECFIN 
                         LEFT JOIN PRODCXC.CCESLPES SALPE 
                         ON SALCO.FICCIDCIAU = SALPE.FICCIDCIAU 
                         AND SALCO.FICCIDPERS = SALPE.FDCCIDPERS 
                         AND SALCO.FICCIDCART = SALPE.FICCIDCART 
                         AND CPERIO.FDCOIDPERI = SALPE.FDCCIDPERI 
                         INNER JOIN 
                         (SELECT ANSAC.FICCIDANCO,ANSAC.FICCIDCIAU, ANSAC.FICCIDSACO, ANSAC.FFCCFECSAC, ANSAC.FDCCIMPORP X1 
                         FROM PRODCXC.CCDANSAD ANSAC 
                         GROUP BY ANSAC.FICCIDCIAU, ANSAC.FICCIDSACO, ANSAC.FFCCFECSAC,ANSAC.FICCIDANCO,ANSAC.FDCCIMPORP) ANTIG 
                         ON SALCO.FICCIDCIAU = ANTIG.FICCIDCIAU 
                         AND SALCO.FICCIDSACO = ANTIG.FICCIDSACO 
                         AND SALCO.FFCCFECSAC = ANTIG.FFCCFECSAC                     
                         WHERE 
                         SALCO.FFCCFECSAC = CPERIO.FFCOFECFIN 
                         AND CPERIO.FICOANIO in (" + anios + ')' + @"      
                         AND SALCO.FICCIDCIAU NOT IN (1,91,92,65,245,15, 58,111,210,121,244) 
                         AND SALCO.FICCIDCIAU IN (" + Id_Agencia + ')' + @"
                         GROUP BY ANTIG.FICCIDANCO,SALCO.FICCIDCIAU, SALCO.FICCIDPERS, SALCO.FICCIDCART, PERSO.FDPEIDTCTE, CIAUN.FSGERAZCOM, CIAUN.FIGEIDMARC, CARTE.FSCCDESCAR, CARTE.FICCIDCART,CPERIO.FICOMES,CPERIO.FICOANIO)    TPRINTER                            
                         ) TP ON TP.FICCIDANCO IN (T2.FSFNNCPT)             
                         GROUP BY T1.FIFNCPTD, T1.FSFNNCPT   
                                        
                         ORDER BY T1.FIFNCPTD";
            #endregion


            return dbCnx.GetDataTable(query);
        }
    }

    public class Reporte
    {
        public int ID_CONCEPTO { get; set; }
        public string CONCEPTO { get; set; }
        public int EXCEL_V1 { get; set; }
        public int WEB_V1 { get; set; }
        public int DIFF_V1 { get; set; }
        public int EXCEL_V1_ACUM { get; set; }
        public int WEB_V1_ACUM { get; set; }
        public int DIFF_V1_ACUM { get; set; }
        public int EXCEL_V2 { get; set; }
        public int WEB_V2 { get; set; }
        public int DIFF_V2 { get; set; }
        public int EXCEL_V2_ACUM { get; set; }
        public int WEB_V2_ACUM { get; set; }
        public int DIFF_V2_ACUM { get; set; }
    }

    public class ReporteVP
    {
        public int ID_CONCEPTO { get; set; }
        public string CONCEPTO { get; set; }
        public int EXCEL_V1 { get; set; }
        public int WEB_V1 { get; set; }
        public int DIFF_V1 { get; set; }
        public int EXCEL_V2 { get; set; }
        public int WEB_V2 { get; set; }
        public int DIFF_V2 { get; set; }
    }

    public class ReporteGlobal
    {
        public int AAZ_ID_CONCEPTO { get; set; }
        public string AAZ_CONCEPTO { get; set; }
        public int AAZ_EXCEL_V1 { get; set; }
        public int AAZ_WEB_V1 { get; set; }
        public int AAZ_DIFF_V1 { get; set; }
        public int AAZ_EXCEL_V2 { get; set; }
        public int AAZ_WEB_V2 { get; set; }
        public int AAZ_DIFF_V2 { get; set; }
        public int AEP_ID_CONCEPTO { get; set; }
        public string AEP_CONCEPTO { get; set; }
        public int AEP_EXCEL_V1 { get; set; }
        public int AEP_WEB_V1 { get; set; }
        public int AEP_DIFF_V1 { get; set; }
        public int AEP_EXCEL_V2 { get; set; }
        public int AEP_WEB_V2 { get; set; }
        public int AEP_DIFF_V2 { get; set; }
        public int ASP_ID_CONCEPTO { get; set; }
        public string ASP_CONCEPTO { get; set; }
        public int ASP_EXCEL_V1 { get; set; }
        public int ASP_WEB_V1 { get; set; }
        public int ASP_DIFF_V1 { get; set; }
        public int ASP_EXCEL_V2 { get; set; }
        public int ASP_WEB_V2 { get; set; }
        public int ASP_DIFF_V2 { get; set; }
    }
}
