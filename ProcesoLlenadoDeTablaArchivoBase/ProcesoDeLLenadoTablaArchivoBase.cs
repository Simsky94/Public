using System;
using System.Collections.Generic;
using DVAModelsReflection;
using DVAModelsReflectionFINA.Models.FINA;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    class ProcesoDeLLenadoTablaArchivoBase
    {
        
        DB2Database _db = null;
        int IdAgencia = 0;
        int anio = 0;
        int mes = 0;
        string Operation = "";

        public ProcesoDeLLenadoTablaArchivoBase(DB2Database _db, int aidAgencia, string Operation, int mes, int anio)
        {
            this._db = _db;
            this.IdAgencia = aidAgencia;
            this.Operation = Operation;
            this.anio = anio;
            this.mes = mes;
            Console.WriteLine("[INICIA Archivo Base]Inicia el proceso de llenado para la agencia ID: " + aidAgencia);
        }
        public void EliminaRegistrosPasadosParaAgencia()
        {
            try
            {
                List<ProcesoLlenado> DataOld;
                if (Operation == "ReplaceAll" || Operation == "SpecificReplaceActual")
                {
                    DataOld = ProcesoLlenado.Listar(_db, mes, anio, IdAgencia);
                }
                else
                {
                    DataOld = ProcesoLlenado.Listar(_db, (mes - 1), anio, IdAgencia);
                }
                Console.WriteLine("[PROCESS] Eliminando " + DataOld.Count + " datos del idMes: " + (mes - 1) + " Año: " + anio + " agencia ID: " + IdAgencia);

                foreach (ProcesoLlenado Data in DataOld)
                {
                    _db.Delete(14091, 999, Data);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[ERROR][EliminaRegistrosPasados]" + " Error: " + e);
            }
        }
        public void InsertaRegistrosActualesParaAgencia()
        {
            try
            {
                DateTime dateIni = DateTime.Now;
                int total = 0;
                List<ProcesoLlenado> Data = ProcesoLlenado.ObtenerData(_db, mes, anio, IdAgencia);
                total = Data.Count;

                //_db.BeginTransaction();
                foreach (ProcesoLlenado item in Data)
                {
                    _db.Insert(14091, 999, item);
                }
                //_db.CommitTransaction();

                DateTime dateFin = DateTime.Now;
                TimeSpan TimeDif = new TimeSpan();
                TimeDif = dateFin - dateIni;
                Console.WriteLine("[FINALIZA Archivo Base][TotalTime=" + TimeDif + "][TotalInsert=" + total + "]Proceso Terminado para la agencia ID: " + IdAgencia);
            }
            catch (Exception e)
            {
                Console.WriteLine("[ERROR][EjecutaParaAgencia][Agencia=" + IdAgencia + "] Error: " + e);
            }
        }
    }
}
