//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using DVAModelsReflection;
//using DVAModelsReflection.Models.GRAL;
//using DVAModelsReflection.Models.FINA;
//using DVADB;
//using DVAModelsReflection.Models.CONT;

//namespace ProcesoLlenadoDeTablaArchivoBase
//{
//    class ProcesoSC
//    {
//        static void Main(string[] args)
//        {
            
//            DB2Database _db = new DB2Database();
//            PeriodoContable periodo  = PeriodoContable.BuscarPorMesAnio(_db, DateTime.Now.Month, DateTime.Now.Year);

//            string ParamOperation = "ALL";
//            int ParamIdAgencia = 0;

//            if (args.Length > 0)
//            {
//                if (args.Length == 1)
//                {
//                    if (args[0] == "")
//                    {
//                        ParamOperation = "ALL";
//                    }else
//                    {
//                        ParamOperation = args[0];
//                    }
//                }
//                else if (args.Length == 2)
//                {
//                    if (!int.TryParse(args[0], out ParamIdAgencia))
//                    {
//                        ParamIdAgencia = 0;
//                    }
//                }
//            }

//            //FUNCION TEMPORAL

//            for (int i = 2018; i < 2019; i++)
//            {
//                Console.WriteLine("Inicia Año: " + i);
//                for (int x = 1; x <= 2 ; x++)
//                {
//                    //if (i == 2017 && x >= 10)
//                    //    return;

//                    Console.WriteLine("Mes: " + x + " Del Año " + i);
//                    List<AgenciasReportes> LiAgenciasReportes = AgenciasReportes.Listar(_db, 1);
//                    Console.WriteLine("Total Agencias: " + LiAgenciasReportes.Count);
//                    foreach (AgenciasReportes aAgenciasReportes in LiAgenciasReportes)
//                    {
//                        try
//                        {
//                            _db.BeginTransaction();
//                            ProcesoSCLlenado proceso = new ProcesoSCLlenado(_db, aAgenciasReportes.IdAgencia, x, i);
//                            proceso.InsertaRegistros();
//                            _db.CommitTransaction();
//                        }
//                        catch (Exception e)
//                        {
//                            Console.WriteLine("[ERROR][Main-All][Agencia=" + aAgenciasReportes.Id + "]" + " Error: " + e);
//                        }
//                    }
//                }
//            }

//        }
//    }

//    public class ProcesoSCLlenado
//    {
//        DB2Database _db = null;
//        int IdAgencia = 0;
//        int anio = 0;
//        int mes = 0;

//        public ProcesoSCLlenado(DB2Database _db, int aidAgencia, int mes, int anio)
//        {
//            this._db = _db;
//            this.IdAgencia = aidAgencia;
//            this.mes = mes;
//            this.anio = anio;
//            Console.WriteLine("[INICIA SI]Inicia el proceso de llenado para la agencia ID: " + aidAgencia);
//        }

//        public void InsertaRegistros()
//        {
//            //Todo aqui
//            try
//            {
//                DateTime dateIni = DateTime.Now;
//                int total = 0;

                
//                List<ProcesoSituacionCartera> RMData =ProcesoSituacionCartera.ListarFromQueryCalculos2(_db,IdAgencia,anio,mes);
//                total = RMData.Count;


//                foreach (ProcesoSituacionCartera item in RMData)
//                {
//                    _db.Insert(14091, 999, item);
//                }
                
                
//                DateTime dateFin = DateTime.Now;
//                TimeSpan TimeDif = new TimeSpan();
//                TimeDif = dateFin - dateIni;
//                Console.WriteLine("[FINALIZA SI][TotalTime=" + TimeDif + "][TotalInsert=" + total + "]Proceso Terminado para la agencia ID: " + IdAgencia);
//            }
//            catch (Exception e)
//            {
//                Console.WriteLine("[ERROR][EjecutaParaAgencia][Agencia=" + IdAgencia + "] Error: " + e);
//            }
//        }
//    }

     
//}
