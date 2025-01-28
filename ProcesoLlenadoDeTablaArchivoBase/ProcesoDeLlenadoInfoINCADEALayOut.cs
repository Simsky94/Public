using DVAModelsReflection;
using DVAModelsReflection.Models.CONT;
using DVAModelsReflectionFINA.Models.FINA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace ProcesoLlenadoDeTablaArchivoBase
{
    class ProcesoDeLlenadoInfoINCADEALayOut
    {
        DB2Database _db = null;
        
        int IdAgencia = 0;
        int anio = 0;
        int mes = 0;
        List<ConceptosContables> liConceptos = new List<ConceptosContables>();
        string ruta = @"C:\Users\jasoria\Desktop\Escritorio\PMO\Proyectos 2022\ARCHIVO BASE V2\Archivos INCADEA\";
        public ProcesoDeLlenadoInfoINCADEALayOut(DB2Database _db, int aAnio, int aMes, int aIdAgencia)
        {
            this._db = _db;
            this.IdAgencia = aIdAgencia;
            this.mes = aMes;
            this.anio = aAnio;
            liConceptos = ConceptosContables.Listar(_db);
            Console.WriteLine("[INICIA RM]Inicia el proceso de llenado para la agencia ID: " + aIdAgencia);
        }

        public void LecturaExcel()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ruta + "Layout_" + anio + "_" + mes + "_" + IdAgencia + ".xlsx");
            //Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            //Excel.Range xlRange = xlWorksheet.UsedRange;

            //int rowCount = xlRange.Rows.Count;
            //int colCount = xlRange.Columns.Count;

            try
            {
                _db = new DB2Database();
                _db.BeginTransaction();

                LeerBalanzaDeComprobacion(ref _db, ref xlWorkbook, anio, mes);

                _db.CommitTransaction();
            }
            catch (Exception ex)
            {
                _db.RollbackTransaction();
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            //Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        public void LeerBalanzaDeComprobacion(ref DB2Database _db, ref Excel.Workbook _xlWorkbook, int anio, int mes)
        {
            Excel._Worksheet xlWorksheet = _xlWorkbook.Sheets[1]; //1_BALANZA_DE_COMPROBACION
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            SaldoPorPeriodoPorCuenta saldo = new SaldoPorPeriodoPorCuenta();
            SaldoPorPeriodoPorCuentaV2 saldoV2 = new SaldoPorPeriodoPorCuentaV2();
            List<SaldoPorPeriodoPorCuenta> saldos = new List<SaldoPorPeriodoPorCuenta>();
            List<SaldoPorPeriodoPorCuentaV2> saldosV2 = new List<SaldoPorPeriodoPorCuentaV2>();
            PeriodoContable periodo = PeriodoContable.BuscarPorMesAnio(_db, mes, anio);
            List<CuentaContable> cuentas = CuentaContable.ListarPorTipoDeCuenta(_db, CuentaContable.ETipo.CATALOGO_INCADEA);
            bool banV2 = false;

            for (int i = 2; i < rowCount; i++)
            {
                saldo = new SaldoPorPeriodoPorCuenta();
                saldo.IdPeriodoContable = periodo.Id;

                saldoV2 = new SaldoPorPeriodoPorCuentaV2();
                saldoV2.IdPeriodoContable = periodo.Id;

                for (int j = 1; j < colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value != null)
                        Console.Write(xlRange.Cells[i, j].Value.ToString() + "\t");
                    else
                        continue;

                    string columna = xlRange.Cells[1, j].Value.ToString();
                    string texto = xlRange.Cells[i, j].Value.ToString();

                    columna = xlRange.Cells[1, 10].Value.ToString();
                    if (xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value != null)
                        texto = xlRange.Cells[i, 10].Value.ToString();
                    else
                        texto = "0";

                    if (texto == "1")
                    {
                        if (!banV2)
                            banV2 = true;
                    }

                    columna = xlRange.Cells[1, j].Value.ToString();
                    texto = xlRange.Cells[i, j].Value.ToString();

                    if (texto == "")
                        continue;

                    if ((columna == "ANIO") || (columna == "MES") || (columna == "DESCRIPCION_CUENTA_CONTABLE"))
                        continue;

                    if (columna == "ID_EMPRESA")
                    {
                        if (banV2)
                            saldoV2.IdAgencia = Convert.ToInt32(texto);
                        else
                            saldo.IdAgencia = Convert.ToInt32(texto);
                    }

                    if (columna == "CUENTA_CONTABLE")
                    {
                        if (banV2) 
                        {
                            saldoV2.TipoDeCuenta = CuentaContable.ETipo.CATALOGO_INCADEA;
                            saldoV2.IdCuentaContable = cuentas.Find(x => x.Clave == texto).Id;
                        }
                        else
                        {
                            saldo.TipoDeCuenta = CuentaContable.ETipo.CATALOGO_INCADEA;
                            saldo.IdCuentaContable = cuentas.Find(x => x.Clave == texto).Id;
                        }
                    }

                    if (columna == "SALDO_INICIAL")
                    {
                        if (banV2)
                        {
                            try
                            {
                                saldoV2.SaldoInicial = Convert.ToDecimal(texto);
                            }
                            catch
                            {
                                saldoV2.SaldoInicial = Convert.ToDecimal(texto.Replace(texto.Substring(texto.IndexOf('E'), texto.Length - texto.IndexOf('E')), ""));
                            }
                        }
                        else
                        {
                            try
                            {
                                saldo.SaldoInicial = Convert.ToDecimal(texto);
                            }
                            catch
                            {
                                saldo.SaldoInicial = Convert.ToDecimal(texto.Replace(texto.Substring(texto.IndexOf('E'), texto.Length - texto.IndexOf('E')), ""));
                            }
                        }
                    }

                    if (columna == "TOTAL_DE_CARGOS")
                    {
                        if (banV2)
                        {
                            try
                            {
                                saldoV2.TotalDeCargos = Convert.ToDecimal(texto);
                            }
                            catch
                            {
                                saldoV2.TotalDeCargos = Convert.ToDecimal(texto.Replace(texto.Substring(texto.IndexOf('E'), texto.Length - texto.IndexOf('E')), ""));
                            }
                        }
                        else
                        {
                            try
                            {
                                saldo.TotalDeCargos = Convert.ToDecimal(texto);
                            }
                            catch
                            {
                                saldo.TotalDeCargos = Convert.ToDecimal(texto.Replace(texto.Substring(texto.IndexOf('E'), texto.Length - texto.IndexOf('E')), ""));
                            }
                        }
                    }

                    if (columna == "TOTAL_DE_ABONOS")
                    {
                        if (banV2)
                        {
                            try
                            {
                                saldoV2.TotalDeAbonos = Convert.ToDecimal(texto);
                            }
                            catch
                            {
                                saldoV2.TotalDeAbonos = Convert.ToDecimal(texto.Replace(texto.Substring(texto.IndexOf('E'), texto.Length - texto.IndexOf('E')), ""));
                            }
                        }
                        else
                        {
                            try
                            {
                                saldo.TotalDeAbonos = Convert.ToDecimal(texto);
                            }
                            catch
                            {
                                saldo.TotalDeAbonos = Convert.ToDecimal(texto.Replace(texto.Substring(texto.IndexOf('E'), texto.Length - texto.IndexOf('E')), ""));
                            }
                        }
                    }

                    if (columna == "SALDO_FINAL")
                    {
                        if (banV2)
                        {
                            try
                            {
                                saldoV2.SaldoFinal = Convert.ToDecimal(texto);
                            }
                            catch
                            {
                                saldoV2.SaldoFinal = Convert.ToDecimal(texto.Replace(texto.Substring(texto.IndexOf('E'), texto.Length - texto.IndexOf('E')), ""));
                            }
                        }
                        else
                        {
                            try
                            {
                                saldo.SaldoFinal = Convert.ToDecimal(texto);
                            }
                            catch
                            {
                                saldo.SaldoFinal = Convert.ToDecimal(texto.Replace(texto.Substring(texto.IndexOf('E'), texto.Length - texto.IndexOf('E')), ""));
                            }
                        }
                    }
                }

                if (banV2)
                    saldosV2.Add(saldoV2);
                else
                    saldos.Add(saldo);

                //banV2 = false;
            }

            if (banV2)
            {
                foreach (SaldoPorPeriodoPorCuentaV2 _saldoV2 in saldosV2)
                {
                    try
                    {
                        _db.Insert(3665, 1241, _saldoV2);
                    }
                    catch (Exception ex)
                    {
                        _db.RollbackTransaction();
                    }
                }
            }
            else
            {
                foreach (SaldoPorPeriodoPorCuenta _saldo in saldos)
                {
                    try
                    {
                        _db.Insert(3665, 1241, _saldo);
                    }
                    catch (Exception ex) 
                    {
                        _db.RollbackTransaction();
                    }
                }
            }

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
        }
    }
}
