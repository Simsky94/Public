using System;
using System.Collections.Generic;
using DVAModelsReflection;
using DVAModelsReflectionFINA.Models.FINA;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    class ProcesoDeLlenadoSC
    {

        DB2Database _db = new DB2Database();
        //PeriodoContable periodo = PeriodoContable.BuscarPorMesAnio(_db, DateTime.Now.Month, DateTime.Now.Year);
            
        
        int IdAgencia = 0;
        int anio = 0;
        int mes = 0;

        public ProcesoDeLlenadoSC(DB2Database _db, int aidAgencia, int mes, int anio)
        {
            this._db = _db;
            this.IdAgencia = aidAgencia;
            this.mes = mes;
            this.anio = anio;
            Console.WriteLine("[INICIA SI]Inicia el proceso de llenado para la agencia ID: " + aidAgencia);
        }

        public void InsertaRegistros()
        {
            //Todo aqui
            try
            {
                DateTime dateIni = DateTime.Now;
                int total = 0;


                List<ProcesoSituacionCartera> RMData = ProcesoSituacionCartera.ListarFromQueryCalculos2(_db, IdAgencia, anio, mes);
                total = RMData.Count;


                foreach (ProcesoSituacionCartera item in RMData)
                {
                    _db.Insert(14091, 999, item);
                }

                DateTime dateFin = DateTime.Now;
                TimeSpan TimeDif = new TimeSpan();
                TimeDif = dateFin - dateIni;
                Console.WriteLine("[FINALIZA SI][TotalTime=" + TimeDif + "][TotalInsert=" + total + "]Proceso Terminado para la agencia ID: " + IdAgencia);
            }
            catch (Exception e)
            {
                Console.WriteLine("[ERROR][EjecutaParaAgencia][Agencia=" + IdAgencia + "] Error: " + e);
            }
        }
    }

}
