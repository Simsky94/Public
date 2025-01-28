using DVAModelsReflection;
using DVAModelsReflectionFINA.Models.FINA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcesoLlenadoDeTablaArchivoBase
{
    public class ProcesoDeLlenadoSI
    {
        DB2Database _db = null;
        int IdAgencia = 0;
        int anio = 0;
        int mes = 0;        
        string query = "";

        public ProcesoDeLlenadoSI(DB2Database _db, int aidAgencia, int mes, int anio)
        {
            this._db = _db;
            this.IdAgencia = aidAgencia;
            this.mes = mes;
            this.anio = anio;            
            Console.WriteLine("[INICIA RM] Inicia el proceso de llenado SI para la agencia ID: " + aidAgencia);
        }
    }
}
