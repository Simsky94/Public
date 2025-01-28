using DVAModelsReflection;
using DVAModelsReflection.Models.AUSA;
using DVAModelsReflection.Models.GRAL;
using DVAModelsReflectionFINA.Models.FINA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    public class ProcesoDeLlenadoGastoCorporativo
    {
        DB2Database _db = null;

        int idAgencia = 0;
        int anio = 0;
        int mes = 0;
        string siglas = "";
        string pestania = "";
        string ruta = @"D:\Users\jasoria\Desktop\Escritorio\PMO\Proyectos 2024\ARCHIVO BASE\Gasto Corporativo\Gasto corporativo Diciembre_2024_LayOut";

        public ProcesoDeLlenadoGastoCorporativo(DB2Database _db, int anio, int mes, string pestania)
        {
            this._db = _db;
            this.anio = anio;
            this.mes = mes;
            this.pestania = pestania;

            Console.WriteLine("[INICIA LECTURA ARCHIVO GASTO CORPORATIVO " + ruta + "]: ");            
            Console.WriteLine("[AÑO]: " + anio);
            Console.WriteLine("[MES]: " + mes);
            Console.WriteLine("[PESTANIA]: " + pestania);
        }

        public void LecturaExcel()
        {
            try
            {
                DVAExcel.ExcelReader eR = new DVAExcel.ExcelReader(ruta);

                DVAExcel.HojaDeTrabajo hojaDeTrabajo = eR.GetHoja(1);

                string valor = "";
                decimal importe = 0;
                int r = 1;
                int c = 1;

                CapturaOPL opl;

                foreach (object celda in hojaDeTrabajo.Celdas)
                {
                    if (celda != null)
                        valor = celda.ToString();

                    if ((c == 2 || c == 7) && r >= 2)
                    {
                        if (c == 2)
                        {
                            siglas = valor;

                            if (siglas == "AFAC")
                                siglas = "AAZ";
                            else if (siglas == "DUC")
                                siglas = "CCND";
                            else if (siglas == "LYM")
                                siglas = "DAML";
                            else if (siglas == "EAH")
                                siglas = "EANH";
                            else if (siglas == "EAP")
                                siglas = "EANP";
                            else if (siglas == "EAZ")
                                siglas = "EANZ";
                            else if (siglas == "LAR")
                                siglas = "LAU";
                            else if (siglas == "LMP")
                                siglas = "LMI";
                            else if (siglas == "MMB")
                                siglas = "MMCB";
                            else if (siglas == "MMC")
                                siglas = "MMCC";
                            else if (siglas == "MMR")
                                siglas = "MMCR";

                            Agencia agencia = Agencia.BuscarPorSigla(_db, siglas);

                            if (agencia != null)
                            {
                                idAgencia = agencia.Id;

                                Console.WriteLine("[CONCEPTO]: " + agencia.Siglas.ToString());
                            }
                            else
                            {
                                idAgencia = 0;
                                c++;
                                continue;
                            }
                        }
                        
                        if (c == 7 && idAgencia != 0)
                        {
                            opl = CapturaOPL.BuscarXAnioXMesXIdGrupoXIdMaestroXIdConcepto(_db, idAgencia, anio.ToString(), mes, 3, 976, 10333);
                                
                            if (opl == null)
                            {
                                importe = Math.Round(Convert.ToDecimal(valor), 0);

                                opl = new CapturaOPL();
                                opl.IdAgencia = idAgencia;
                                opl.Anio = anio.ToString();
                                opl.IdMes = mes;
                                opl.IdGrupo = 3;
                                opl.IdMaestro = 976;
                                opl.IdConcepto = 10333;
                                opl.Valor = importe;

                                _db.Insert<CapturaOPL>(3665, 0, opl);

                                Console.WriteLine("[ID_AGENCIA]: " + opl.IdAgencia.ToString() + " ==> IMPORTE: " + opl.Valor);
                            }
                        }

                        //if (c == 9 && idAgencia != 0)
                        //{
                        //    importe = Math.Round(Convert.ToDecimal(valor),0);

                        //    opl = new CapturaOPL();
                        //    opl.IdAgencia = idAgencia;
                        //    opl.Anio = anio.ToString();
                        //    opl.IdMes = mes;
                        //    opl.IdGrupo = 3;
                        //    opl.IdMaestro = 976;
                        //    opl.IdConcepto = 10333;
                        //    opl.Valor = importe;

                        //    //_db.Insert<CapturaOPL>(3665, 0, opl);

                        //    Console.WriteLine("[ID_AGENCIA]: " + opl.IdAgencia.ToString() + " ==> IMPORTE: " + opl.Valor);
                        //}
                    }

                    c++;

                    if (r == hojaDeTrabajo.LastRow)
                        r = 1;

                    if (c == hojaDeTrabajo.LastColumn + 2)
                    {
                        c = 1;
                        r++;
                    }
                }

                eR.Dispose();
            }
            catch (Exception ex)
            {
                
            }
        }
    }
}
